import json
import re
from io import BytesIO
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List
import zipfile

import pandas as pd
import streamlit as st


st.set_page_config(page_title="KAU Relabeling Priorisierung", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)


def apply_custom_style() -> None:
    st.markdown(
        """
        <style>
        .stApp {
            background: linear-gradient(160deg, #061739 0%, #0a2a5f 60%, #123d7a 100%);
            color: #e8efff;
        }
        [data-testid="stSidebar"] {
            background: #081a3f;
        }
        h1, h2, h3, h4, h5, h6, p, label, span, div {
            color: #e8efff;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False)
def load_csv_from_path(path_str: str) -> pd.DataFrame:
    return pd.read_csv(path_str)


def list_local_csvs() -> List[Path]:
    return sorted(DATA_DIR.glob("*.csv"), key=lambda p: p.stat().st_mtime, reverse=True)


def get_text_columns(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=["object", "string"]).columns.tolist()


def get_numeric_columns(df: pd.DataFrame) -> List[str]:
    return df.select_dtypes(include=["number"]).columns.tolist()


def short_label(name: str, max_len: int = 24) -> str:
    if len(name) <= max_len:
        return name
    return name[: max_len - 3] + "..."


def infer_col_width(series: pd.Series, name: str) -> str:
    if pd.api.types.is_bool_dtype(series):
        return "small"
    if pd.api.types.is_numeric_dtype(series):
        return "small"
    sample = series.dropna().astype(str).head(80)
    max_len = max([len(name)] + ([sample.str.len().max()] if not sample.empty else [0]))
    if max_len <= 14:
        return "small"
    if max_len <= 35:
        return "medium"
    return "large"


def build_column_config(df: pd.DataFrame, allow_manual_edit: bool) -> Dict[str, Any]:
    config: Dict[str, Any] = {}
    for col in df.columns:
        label = short_label(col)
        width = infer_col_width(df[col], col)
        if col == "prio_manual":
            config[col] = st.column_config.SelectboxColumn(
                label=label,
                options=["", "Low", "Medium", "High"],
                help=col,
                width=width,
            )
        elif col == "prio_note":
            config[col] = st.column_config.TextColumn(label=label, help=col, width="large")
        elif pd.api.types.is_bool_dtype(df[col]):
            config[col] = st.column_config.CheckboxColumn(label=label, help=col, width=width)
        elif pd.api.types.is_numeric_dtype(df[col]):
            config[col] = st.column_config.NumberColumn(label=label, help=col, width=width)
        else:
            config[col] = st.column_config.TextColumn(label=label, help=col, width=width)
    return config


def empty_value_mask(series: pd.Series) -> pd.Series:
    s = series.astype("string")
    normalized = s.str.strip().str.lower()
    placeholders = {"", "none", "null", "nan", "na", "n/a"}
    return series.isna() | normalized.isin(placeholders)


def json_safe(value: Any) -> Any:
    if isinstance(value, (date, datetime)):
        return value.isoformat()
    if isinstance(value, tuple):
        return [json_safe(v) for v in value]
    if isinstance(value, list):
        return [json_safe(v) for v in value]
    if isinstance(value, dict):
        return {k: json_safe(v) for k, v in value.items()}
    return value


def is_forbidden_widget_key(key: str) -> bool:
    return (
        "_preset_" in key
        or "_editor" in key
        or key.endswith("_download")
        or key.endswith("_activate_sidebar")
        or key.endswith("_pending_preset_state")
        or key.endswith("_preset_loaded_msg")
    )


def get_store_key(prefix: str) -> str:
    return f"{prefix}__filter_store"


def get_filter_store(prefix: str) -> Dict[str, Any]:
    return st.session_state.get(get_store_key(prefix), {})


def set_filter_store(prefix: str, state: Dict[str, Any]) -> None:
    cleaned = {
        key: value
        for key, value in state.items()
        if key.startswith(f"{prefix}_") and not is_forbidden_widget_key(key)
    }
    st.session_state[get_store_key(prefix)] = json_safe(cleaned)


def sanitize_filter_stores() -> None:
    for prefix in ["view_a", "view_b", "view_c"]:
        store_key = get_store_key(prefix)
        if store_key in st.session_state:
            set_filter_store(prefix, st.session_state[store_key])


def collect_view_state(prefix: str) -> Dict[str, Any]:
    stored = get_filter_store(prefix)
    if stored:
        return stored
    state = {}
    for key, value in st.session_state.items():
        if key.startswith(f"{prefix}_"):
            if is_forbidden_widget_key(key):
                continue
            state[key] = json_safe(value)
    return state


def collect_state_by_prefix(prefix: str) -> Dict[str, Any]:
    state: Dict[str, Any] = {}
    for key, value in st.session_state.items():
        if key.startswith(prefix):
            state[key] = json_safe(value)
    return state


def build_export_zip(csv_bytes: bytes, csv_name: str, filter_state: Dict[str, Any], state_name: str) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(csv_name, csv_bytes)
        zf.writestr(state_name, json.dumps(filter_state, ensure_ascii=False, indent=2))
    return buf.getvalue()


def remap_state_prefix(state: Dict[str, Any], target_prefix: str) -> Dict[str, Any]:
    source_prefix = None
    for key in state.keys():
        if key.startswith("view_a_"):
            source_prefix = "view_a"
            break
        if key.startswith("view_b_"):
            source_prefix = "view_b"
            break
        if key.startswith("view_c_"):
            source_prefix = "view_c"
            break
    if not source_prefix or source_prefix == target_prefix:
        return state

    remapped = {}
    for key, value in state.items():
        if key.startswith(f"{source_prefix}_"):
            remapped[target_prefix + key[len(source_prefix):]] = value
        else:
            remapped[key] = value
    return remapped


def apply_view_state(state: Dict[str, Any]) -> None:
    clean_state: Dict[str, Any] = {}
    for key, value in state.items():
        if is_forbidden_widget_key(key):
            continue
        if key.endswith("_date_range") and isinstance(value, list) and len(value) == 2:
            try:
                clean_state[key] = (date.fromisoformat(value[0]), date.fromisoformat(value[1]))
            except Exception:
                pass
        elif "_num_rng_" in key and isinstance(value, list) and len(value) == 2:
            try:
                clean_state[key] = (float(value[0]), float(value[1]))
            except Exception:
                pass
        elif key.endswith("_num_cols") and value is None:
            clean_state[key] = []
        else:
            clean_state[key] = value

    # Persist independently of widget lifecycle and seed session_state.
    prefixes = {k.split("_", 1)[0] + "_" + k.split("_", 2)[1] for k in clean_state.keys() if "_" in k}
    for pfx in ["view_a", "view_b", "view_c"]:
        pfx_state = {k: v for k, v in clean_state.items() if k.startswith(f"{pfx}_")}
        if pfx_state:
            current = get_filter_store(pfx)
            current.update(json_safe(pfx_state))
            set_filter_store(pfx, current)
    for key, value in clean_state.items():
        st.session_state[key] = value


def restore_widgets_from_store(prefix: str) -> None:
    store = get_filter_store(prefix)
    if not store:
        return
    for key, value in store.items():
        if key.startswith(f"{prefix}_") and not is_forbidden_widget_key(key):
            if key not in st.session_state:
                st.session_state[key] = value


def sync_store_from_session(prefix: str) -> None:
    store = get_filter_store(prefix).copy()
    for key, value in st.session_state.items():
        if key.startswith(f"{prefix}_") and not is_forbidden_widget_key(key):
            store[key] = json_safe(value)
    set_filter_store(prefix, store)


def _apply_filters_from_state(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    state = get_filter_store(prefix)

    def sget(key: str, default=None):
        return state.get(key, st.session_state.get(key, default))

    filtered = df.copy()
    text_cols = get_text_columns(filtered)

    selected_col = sget(f"{prefix}_regex_col")
    include_regex = sget(f"{prefix}_regex_include", "")
    exclude_regex = sget(f"{prefix}_regex_exclude", "")
    include_empty_with_regex = bool(sget(f"{prefix}_regex_inc_empty", False))
    if selected_col in filtered.columns and include_regex:
        try:
            re.compile(include_regex)
            selected_series = filtered[selected_col].astype("string")
            regex_mask = selected_series.str.contains(include_regex, case=False, regex=True, na=False)
            empty_mask = empty_value_mask(filtered[selected_col])
            combined_mask = regex_mask | empty_mask if include_empty_with_regex else regex_mask
            filtered = filtered[combined_mask]
        except re.error:
            pass

    if selected_col in filtered.columns and not include_empty_with_regex:
        filtered = filtered[~empty_value_mask(filtered[selected_col])]
    if selected_col in filtered.columns and exclude_regex:
        try:
            re.compile(exclude_regex)
            filtered = filtered[
                ~filtered[selected_col].astype(str).str.contains(exclude_regex, case=False, regex=True, na=False)
            ]
        except re.error:
            pass

    for col in text_cols:
        selected_values = sget(f"{prefix}_cat_{col}", [])
        if selected_values:
            include_empty = "(leer)" in selected_values
            selected_non_empty = [v for v in selected_values if v != "(leer)"]
            mask_non_empty = filtered[col].astype(str).isin(selected_non_empty) if selected_non_empty else False
            mask_empty = empty_value_mask(filtered[col])
            if include_empty and selected_non_empty:
                filtered = filtered[mask_non_empty | mask_empty]
            elif include_empty:
                filtered = filtered[mask_empty]
            else:
                filtered = filtered[mask_non_empty]

    selected_numeric = sget(f"{prefix}_num_cols", [])
    for col in selected_numeric:
        if col not in filtered.columns:
            continue
        col_num = pd.to_numeric(filtered[col], errors="coerce")
        if not col_num.notna().any():
            continue
        rng = sget(f"{prefix}_num_rng_{col}", (float(col_num.min()), float(col_num.max())))
        filtered = filtered[col_num.between(rng[0], rng[1])]

    date_col = sget(f"{prefix}_date_col", "(kein Filter)")
    if date_col in df.columns:
        parsed = pd.to_datetime(df[date_col], errors="coerce")
        dr = sget(f"{prefix}_date_range")
        if isinstance(dr, tuple) and len(dr) == 2:
            filtered = filtered[parsed.dt.date.between(dr[0], dr[1]).fillna(False)]

    enabled = sget(f"{prefix}_rule_on", True)
    if enabled:
        result = filtered.copy()
        result["prio_score"] = 0
        result["prio_labels"] = ""
        rule_count = int(sget(f"{prefix}_rule_count", 2))
        for idx in range(rule_count):
            label = sget(f"{prefix}_rule_label_{idx}", f"R{idx+1}")
            col = sget(f"{prefix}_rule_col_{idx}")
            rule_type = sget(f"{prefix}_rule_type_{idx}", "Regex")
            pattern = sget(f"{prefix}_rule_rx_{idx}", "")
            threshold = sget(f"{prefix}_rule_threshold_{idx}", 0.0)
            score = int(sget(f"{prefix}_rule_sc_{idx}", 1))
            if col not in result.columns:
                continue

            if rule_type == "Regex":
                if not pattern:
                    continue
                try:
                    re.compile(pattern)
                    mask = result[col].astype(str).str.contains(pattern, case=False, regex=True, na=False)
                except re.error:
                    continue
            else:
                col_num = pd.to_numeric(result[col], errors="coerce")
                mask = col_num > float(threshold)

            result.loc[mask, "prio_score"] += score
            result.loc[mask, "prio_labels"] = result.loc[mask, "prio_labels"] + ";" + str(label)
            result[f"rule_{idx+1}_match"] = mask
        result["prio_labels"] = result["prio_labels"].str.strip(";")
        result["prio_match"] = result["prio_score"] > 0
        filtered = result

    return filtered


def render_view_preset_tools(prefix: str, view_name: str) -> None:
    c1, c2 = st.columns([2, 1])
    with c1:
        st.text_input("Preset-Name", key=f"{prefix}_preset_name", placeholder=f"{view_name} Fokus")
    with c2:
        st.text_input("Bemerkung", key=f"{prefix}_preset_note", placeholder="Kurznotiz")

    payload = {
        "view": view_name,
        "preset_name": st.session_state.get(f"{prefix}_preset_name", ""),
        "preset_note": st.session_state.get(f"{prefix}_preset_note", ""),
        "state": collect_view_state(prefix),
    }

    c3, c4 = st.columns([1, 1])
    with c3:
        st.download_button(
            "Filteransicht speichern (JSON)",
            data=json.dumps(payload, ensure_ascii=False, indent=2),
            file_name=(st.session_state.get(f"{prefix}_preset_name") or f"{prefix}_preset").replace(" ", "_") + ".json",
            mime="application/json",
            key=f"{prefix}_preset_download",
        )
    with c4:
        uploaded = st.file_uploader("Filteransicht laden (JSON)", type=["json"], key=f"{prefix}_preset_upload")
        if uploaded is not None and st.button("Preset jetzt laden", key=f"{prefix}_preset_apply_btn"):
            try:
                parsed = json.load(uploaded)
                incoming_state = remap_state_prefix(parsed.get("state", {}), prefix)
                st.session_state[f"{prefix}_pending_preset_state"] = incoming_state
                st.session_state[f"{prefix}_preset_loaded_msg"] = True
                st.rerun()
            except Exception as exc:
                st.error(f"Preset konnte nicht geladen werden: {exc}")


def apply_filters(df: pd.DataFrame, prefix: str, show_sidebar_filters: bool) -> pd.DataFrame:
    if not show_sidebar_filters:
        return _apply_filters_from_state(df, prefix)

    restore_widgets_from_store(prefix)
    text_cols = get_text_columns(df)
    numeric_cols = get_numeric_columns(df)
    filtered = df.copy()

    st.sidebar.markdown("### Filter")
    with st.sidebar.expander("Regex-Filter", expanded=True) as regex_box:
        if text_cols:
            selected_col = regex_box.selectbox("Regex-Spalte", text_cols, key=f"{prefix}_regex_col")
            include_regex = regex_box.text_input("Regex Include", key=f"{prefix}_regex_include")
            exclude_regex = regex_box.text_input("Regex Exclude", key=f"{prefix}_regex_exclude")
            include_empty_with_regex = regex_box.checkbox(
                "Leere bei Include einschließen",
                value=False,
                key=f"{prefix}_regex_inc_empty",
            )

            if include_regex:
                try:
                    re.compile(include_regex)
                    selected_series = filtered[selected_col].astype("string")
                    regex_mask = selected_series.str.contains(include_regex, case=False, regex=True, na=False)
                    empty_mask = empty_value_mask(filtered[selected_col])
                    combined_mask = regex_mask | empty_mask if include_empty_with_regex else regex_mask
                    filtered = filtered[combined_mask]
                except re.error as exc:
                    st.sidebar.error(f"Ungültiger Include-Regex: {exc}")

            if not include_empty_with_regex:
                filtered = filtered[~empty_value_mask(filtered[selected_col])]

            if exclude_regex:
                try:
                    re.compile(exclude_regex)
                    filtered = filtered[
                        ~filtered[selected_col].astype(str).str.contains(exclude_regex, case=False, regex=True, na=False)
                    ]
                except re.error as exc:
                    st.sidebar.error(f"Ungültiger Exclude-Regex: {exc}")

    with st.sidebar.expander("Kategorienfilter", expanded=False) as cat_box:
        for col in text_cols:
            col_str = filtered[col].astype("string")
            non_empty_values = sorted(
                col_str[
                    ~empty_value_mask(filtered[col])
                ].unique().tolist()
            )
            options = non_empty_values + (["(leer)"] if empty_value_mask(filtered[col]).any() else [])
            if 1 < len(options) <= 101:
                selected_values = cat_box.multiselect(col, options, key=f"{prefix}_cat_{col}")
                if selected_values:
                    include_empty = "(leer)" in selected_values
                    selected_non_empty = [v for v in selected_values if v != "(leer)"]
                    mask_non_empty = filtered[col].astype(str).isin(selected_non_empty) if selected_non_empty else False
                    mask_empty = empty_value_mask(filtered[col])
                    if include_empty and selected_non_empty:
                        filtered = filtered[mask_non_empty | mask_empty]
                    elif include_empty:
                        filtered = filtered[mask_empty]
                    else:
                        filtered = filtered[mask_non_empty]

    with st.sidebar.expander("Numerische Filter", expanded=False) as num_box:
        selected_numeric = num_box.multiselect("Spalten auswählen", numeric_cols, key=f"{prefix}_num_cols")
        numeric_ranges: Dict[str, tuple] = {}
        for col in selected_numeric:
            # Keep each slider independent by deriving bounds from original df.
            col_num = pd.to_numeric(df[col], errors="coerce")
            if not col_num.notna().any():
                continue
            cmin = float(col_num.min())
            cmax = float(col_num.max())
            if cmin == cmax:
                continue
            slider_key = f"{prefix}_num_rng_{col}"
            default_rng = st.session_state.get(slider_key, (cmin, cmax))
            if not isinstance(default_rng, tuple) or len(default_rng) != 2:
                default_rng = (cmin, cmax)
            imin = int(round(cmin))
            imax = int(round(cmax))
            idef = (
                max(imin, int(round(float(default_rng[0])))),
                min(imax, int(round(float(default_rng[1])))),
            )
            if idef[0] > idef[1]:
                idef = (imin, imax)
            rng = num_box.slider(
                f"{col} Bereich",
                min_value=imin,
                max_value=imax,
                value=idef,
                step=1,
                key=slider_key,
            )
            numeric_ranges[col] = rng

        for col, rng in numeric_ranges.items():
            col_num_filtered = pd.to_numeric(filtered[col], errors="coerce")
            filtered = filtered[col_num_filtered.between(rng[0], rng[1])]

    with st.sidebar.expander("Datumsfilter", expanded=False) as date_box:
        date_col = date_box.selectbox("Datumsspalte", ["(kein Filter)"] + list(df.columns), key=f"{prefix}_date_col")
        if date_col != "(kein Filter)":
            parsed = pd.to_datetime(df[date_col], errors="coerce")
            valid = parsed.dropna()
            if not valid.empty:
                dmin = valid.min().date()
                dmax = valid.max().date()
                date_key = f"{prefix}_date_range"
                default_dr = st.session_state.get(date_key, (dmin, dmax))
                if not isinstance(default_dr, tuple) or len(default_dr) != 2:
                    default_dr = (dmin, dmax)
                start_d = default_dr[0] if isinstance(default_dr[0], date) else dmin
                end_d = default_dr[1] if isinstance(default_dr[1], date) else dmax
                if start_d < dmin:
                    start_d = dmin
                if end_d > dmax:
                    end_d = dmax
                if start_d > end_d:
                    start_d, end_d = dmin, dmax
                dr = date_box.date_input(
                    "Zeitraum",
                    value=(start_d, end_d),
                    min_value=dmin,
                    max_value=dmax,
                    key=date_key,
                )
                if isinstance(dr, tuple) and len(dr) == 2:
                    filtered = filtered[parsed.dt.date.between(dr[0], dr[1]).fillna(False)]

    with st.sidebar.expander("Rule Engine", expanded=False) as rule_box:
        enabled = rule_box.checkbox("Regeln aktivieren", value=True, key=f"{prefix}_rule_on")
        if enabled:
            result = filtered.copy()
            result["prio_score"] = 0
            result["prio_labels"] = ""
            rule_count = rule_box.slider("Anzahl Regeln", 1, 5, 2, key=f"{prefix}_rule_count")
            rule_cols = get_text_columns(result)
            for idx in range(rule_count):
                label = rule_box.text_input(f"Label {idx+1}", value=f"R{idx+1}", key=f"{prefix}_rule_label_{idx}")
                rule_type = rule_box.selectbox(
                    f"Regeltyp {idx+1}",
                    ["Regex", "Numerisch >"],
                    key=f"{prefix}_rule_type_{idx}",
                )
                if rule_type == "Regex":
                    col = rule_box.selectbox(
                        f"Spalte {idx+1}",
                        rule_cols if rule_cols else list(result.columns),
                        key=f"{prefix}_rule_col_{idx}",
                    )
                    pattern = rule_box.text_input(f"Regex {idx+1}", key=f"{prefix}_rule_rx_{idx}")
                    threshold = None
                else:
                    numeric_cols_rule = get_numeric_columns(result)
                    if numeric_cols_rule:
                        col = rule_box.selectbox(
                            f"Spalte {idx+1}",
                            numeric_cols_rule,
                            key=f"{prefix}_rule_col_{idx}",
                        )
                    else:
                        col = None
                        rule_box.caption("Keine numerischen Spalten verfügbar.")
                    threshold = rule_box.number_input(
                        f"Schwellwert {idx+1}",
                        value=0.0,
                        key=f"{prefix}_rule_threshold_{idx}",
                    )
                    pattern = ""
                score = rule_box.number_input(
                    f"Score {idx+1}",
                    min_value=-10,
                    max_value=1000,
                    value=1,
                    key=f"{prefix}_rule_sc_{idx}",
                )
                if col is None:
                    continue

                if rule_type == "Regex":
                    if pattern:
                        try:
                            re.compile(pattern)
                            mask = result[col].astype(str).str.contains(pattern, case=False, regex=True, na=False)
                            result.loc[mask, "prio_score"] += int(score)
                            result.loc[mask, "prio_labels"] = result.loc[mask, "prio_labels"] + ";" + label
                            result[f"rule_{idx+1}_match"] = mask
                        except re.error as exc:
                            st.sidebar.error(f"Regel {idx+1}: {exc}")
                else:
                    col_num = pd.to_numeric(result[col], errors="coerce")
                    mask = col_num > float(threshold if threshold is not None else 0.0)
                    result.loc[mask, "prio_score"] += int(score)
                    result.loc[mask, "prio_labels"] = result.loc[mask, "prio_labels"] + ";" + label
                    result[f"rule_{idx+1}_match"] = mask
            result["prio_labels"] = result["prio_labels"].str.strip(";")
            result["prio_match"] = result["prio_score"] > 0
            filtered = result

    sync_store_from_session(prefix)
    return filtered


def apply_manual_override(df: pd.DataFrame, prefix: str) -> pd.DataFrame:
    edited = df.copy()
    if "prio_score" not in edited.columns:
        edited["prio_score"] = 0
        edited["prio_labels"] = ""
        edited["prio_match"] = False

    if "prio_manual" not in edited.columns:
        edited["prio_manual"] = ""
    if "prio_note" not in edited.columns:
        edited["prio_note"] = ""

    st.markdown("#### Editierbare Ansicht")
    manual_on = st.checkbox("Manuellen Override aktivieren", value=True, key=f"{prefix}_manual_on")

    column_config = build_column_config(edited, allow_manual_edit=manual_on)

    edited = st.data_editor(
        edited,
        use_container_width=True,
        height=420,
        key=f"{prefix}_editor",
        column_config=column_config,
        disabled=[c for c in edited.columns if (c not in ["prio_manual", "prio_note"] or not manual_on)],
    )

    manual_score_map = {"Low": 1, "Medium": 5, "High": 10}
    edited["prio_final_score"] = edited.apply(
        lambda r: manual_score_map.get(str(r.get("prio_manual", "")).strip(), r.get("prio_score", 0)),
        axis=1,
    )
    edited["prio_final_label"] = edited.apply(
        lambda r: str(r.get("prio_manual", "")).strip() if str(r.get("prio_manual", "")).strip() else str(r.get("prio_labels", "")),
        axis=1,
    )
    edited["prio_final_match"] = pd.to_numeric(edited["prio_final_score"], errors="coerce").fillna(0) > 0
    return edited


def render_summary(df: pd.DataFrame, prefix: str) -> None:
    st.markdown("#### Summenanzeigen")
    text_cols = get_text_columns(df)
    if not text_cols:
        return
    default_group = "Bereich" if "Bereich" in df.columns else text_cols[0]
    group_col = st.selectbox("Gruppieren nach", text_cols, index=text_cols.index(default_group), key=f"{prefix}_group_col")

    summary = df.groupby(group_col, dropna=False).size().reset_index(name="anzahl_messstellen")
    summary = summary.sort_values("anzahl_messstellen", ascending=False)

    match_col = "prio_final_match" if "prio_final_match" in df.columns else "prio_match"
    score_col = "prio_final_score" if "prio_final_score" in df.columns else "prio_score"

    if match_col in df.columns:
        hits = df[df[match_col]].groupby(group_col, dropna=False).size().reset_index(name="anzahl_prio_treffer")
        summary = summary.merge(hits, on=group_col, how="left")
        summary["anzahl_prio_treffer"] = summary["anzahl_prio_treffer"].fillna(0).astype(int)

    if score_col in df.columns:
        score_sum = pd.to_numeric(df[score_col], errors="coerce").groupby(df[group_col], dropna=False).sum().reset_index(name="summe_prio_score")
        summary = summary.merge(score_sum, on=group_col, how="left")

    st.dataframe(
        summary,
        use_container_width=True,
        height=260,
        column_config=build_column_config(summary, allow_manual_edit=False),
    )


def render_view(df: pd.DataFrame, view_name: str, prefix: str, active_prefix: str) -> None:
    st.markdown(f"### {view_name}")
    if st.button("Diese Ansicht links bearbeiten", key=f"{prefix}_activate_sidebar"):
        st.session_state["active_prefix"] = prefix
        st.rerun()
    if active_prefix != prefix:
        st.caption("Hinweis: Sidebar-Filter sind aktuell auf eine andere Ansicht gesetzt.")
    if st.session_state.pop(f"{prefix}_preset_loaded_msg", False):
        st.success("Preset geladen. Ansicht aktualisiert.")

    filtered = apply_filters(df, prefix, show_sidebar_filters=(prefix == active_prefix))
    render_view_preset_tools(prefix, view_name)
    st.write(f"{len(filtered)} von {len(df)} Zeilen")
    edited = apply_manual_override(filtered, prefix)
    render_summary(edited, prefix)

    csv_name = f"relabeling_{prefix}.csv"
    csv_bytes = edited.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        f"{view_name}: CSV exportieren",
        csv_bytes,
        file_name=csv_name,
        mime="text/csv",
        key=f"{prefix}_download",
    )
    export_state = {
        "view": view_name,
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "filters": collect_view_state(prefix),
    }
    zip_bytes = build_export_zip(
        csv_bytes=csv_bytes,
        csv_name=csv_name,
        filter_state=export_state,
        state_name=f"{prefix}_filter_state.json",
    )
    st.download_button(
        f"{view_name}: CSV + Filter (ZIP)",
        zip_bytes,
        file_name=f"relabeling_{prefix}_with_filters.zip",
        mime="application/zip",
        key=f"{prefix}_download_zip",
    )


def render_standort_analyse(df: pd.DataFrame) -> None:
    st.markdown("### Standort-Analyse")

    text_cols = get_text_columns(df)
    if not text_cols:
        st.info("Keine Textspalten verfügbar.")
        return

    default_zug = "Zugänglichkeit" if "Zugänglichkeit" in df.columns else text_cols[0]
    default_standort = "Standort" if "Standort" in df.columns else text_cols[0]

    c1, c2 = st.columns([1, 1])
    zug_col = c1.selectbox("Spalte Zugänglichkeit", text_cols, index=text_cols.index(default_zug), key="sa_zug_col")
    standort_col = c2.selectbox("Spalte Standort", text_cols, index=text_cols.index(default_standort), key="sa_standort_col")

    c3, c4, c5 = st.columns([2, 2, 1])
    include_regex = c3.text_input("Regex Include (Zugänglichkeit)", key="sa_inc_rx")
    exclude_regex = c4.text_input("Regex Exclude (Zugänglichkeit)", key="sa_exc_rx")
    include_empty = c5.checkbox("Leer einschließen", value=False, key="sa_inc_empty")

    filtered = df.copy()
    zug_series = filtered[zug_col].astype("string")
    empty_mask = filtered[zug_col].isna() | (zug_series.str.strip() == "")
    include_mask = pd.Series([True] * len(filtered), index=filtered.index)

    if include_regex:
        try:
            re.compile(include_regex)
            regex_mask = zug_series.str.contains(include_regex, case=False, regex=True, na=False)
            include_mask = regex_mask | empty_mask if include_empty else regex_mask
        except re.error as exc:
            st.error(f"Ungültiger Include-Regex: {exc}")
            include_mask = (~empty_mask) if not include_empty else pd.Series([True] * len(filtered), index=filtered.index)
    else:
        include_mask = pd.Series([True] * len(filtered), index=filtered.index) if include_empty else (~empty_mask)

    filtered = filtered[include_mask]
    zug_series = filtered[zug_col].astype("string")

    if exclude_regex:
        try:
            re.compile(exclude_regex)
            filtered = filtered[~zug_series.str.contains(exclude_regex, case=False, regex=True, na=False)]
        except re.error as exc:
            st.error(f"Ungültiger Exclude-Regex: {exc}")

    c6, c7, c8 = st.columns([2, 2, 1])
    basis_mode = c6.selectbox(
        "Standort-Basis",
        ["Original", "Vor erstem Leerzeichen", "Vor Trennzeichen"],
        key="sa_basis_mode",
    )
    separator = c7.text_input("Trennzeichen", value=" ", key="sa_sep")
    starts_with = c8.text_input("Standort beginnt mit", key="sa_prefix")

    standort_series = filtered[standort_col].astype("string")
    if starts_with:
        filtered = filtered[standort_series.str.startswith(starts_with, na=False)]
        standort_series = filtered[standort_col].astype("string")

    if basis_mode == "Original":
        filtered["standort_basis"] = standort_series.fillna("")
    elif basis_mode == "Vor erstem Leerzeichen":
        filtered["standort_basis"] = standort_series.fillna("").str.split().str[0].fillna("")
    else:
        sep = separator if separator else " "
        filtered["standort_basis"] = standort_series.fillna("").str.split(sep).str[0].fillna("")

    filtered = filtered[filtered["standort_basis"].astype(str).str.strip() != ""]

    st.write(f"{len(filtered)} von {len(df)} Zeilen")

    summary = (
        filtered.groupby("standort_basis", dropna=False)
        .agg(
            anzahl_messstellen=(standort_col, "size"),
            varianten_standort=(standort_col, pd.Series.nunique),
        )
        .reset_index()
        .sort_values("anzahl_messstellen", ascending=False)
    )
    st.dataframe(
        summary,
        use_container_width=True,
        height=300,
        column_config=build_column_config(summary, allow_manual_edit=False),
    )

    basis_options = summary["standort_basis"].astype(str).tolist()
    if basis_options:
        c9, c10 = st.columns([2, 2])
        selected_basis = c9.selectbox("Details für Standort-Basis", basis_options, key="sa_basis_pick")
        detail_search = c10.text_input("Details einschränken (Standort enthält)", key="sa_detail_search")
        detail = (
            filtered[filtered["standort_basis"].astype(str) == selected_basis]
            .groupby(standort_col, dropna=False)
            .size()
            .reset_index(name="anzahl_messstellen")
            .sort_values("anzahl_messstellen", ascending=False)
        )
        if detail_search:
            detail = detail[detail[standort_col].astype("string").str.contains(detail_search, case=False, na=False)]
        st.dataframe(
            detail,
            use_container_width=True,
            height=260,
            column_config=build_column_config(detail, allow_manual_edit=False),
        )

        st.markdown("#### Rohdaten für gewählte Standort-Basis")
        raw = filtered[filtered["standort_basis"].astype(str) == selected_basis].copy()
        if detail_search:
            raw = raw[raw[standort_col].astype("string").str.contains(detail_search, case=False, na=False)]
        show_cols_default = [col for col in [standort_col, zug_col, "standort_basis"] if col in raw.columns]
        show_cols = st.multiselect(
            "Spalten anzeigen",
            options=list(raw.columns),
            default=show_cols_default if show_cols_default else list(raw.columns)[: min(8, len(raw.columns))],
            key="sa_raw_cols",
        )
        raw_view = raw[show_cols] if show_cols else raw
        st.dataframe(
            raw_view,
            use_container_width=True,
            height=320,
            column_config=build_column_config(raw_view, allow_manual_edit=False),
        )

    sa_csv_name = "standort_analyse.csv"
    sa_csv_bytes = filtered.to_csv(index=False).encode("utf-8-sig")
    st.download_button(
        "Standort-Analyse exportieren (CSV)",
        sa_csv_bytes,
        file_name=sa_csv_name,
        mime="text/csv",
        key="sa_export",
    )
    sa_state = {
        "view": "Standort-Analyse",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "filters": collect_state_by_prefix("sa_"),
    }
    sa_zip_bytes = build_export_zip(
        csv_bytes=sa_csv_bytes,
        csv_name=sa_csv_name,
        filter_state=sa_state,
        state_name="standort_analyse_filter_state.json",
    )
    st.download_button(
        "Standort-Analyse: CSV + Filter (ZIP)",
        sa_zip_bytes,
        file_name="standort_analyse_with_filters.zip",
        mime="application/zip",
        key="sa_export_zip",
    )


def main() -> None:
    apply_custom_style()
    sanitize_filter_stores()
    st.title("KAU Relabeling Priorisierung")

    st.sidebar.header("Datenquelle")
    with st.sidebar.expander("CSV Auswahl / Upload", expanded=False):
        csv_files = list_local_csvs()
        selected_path = None
        if csv_files:
            options = {p.name: str(p) for p in csv_files}
            selected_name = st.selectbox("Aktuelle CSV", list(options.keys()))
            selected_path = options[selected_name]
            st.caption(f"Pfad: {selected_path}")
        else:
            st.info("Noch keine lokale CSV in /data vorhanden.")

        uploaded_file = st.file_uploader("Neue CSV hochladen", type=["csv"], key="csv_upload")
        if uploaded_file is not None:
            keep_uploaded = st.checkbox("Hochgeladene CSV im Projekt speichern", value=True)
            if keep_uploaded:
                filename = st.text_input("Dateiname", value=uploaded_file.name)
                if st.button("CSV speichern und verwenden"):
                    target = DATA_DIR / Path(filename).name
                    target.write_bytes(uploaded_file.getvalue())
                    st.success(f"Gespeichert: {target}")
                    st.rerun()
            else:
                temp_path = DATA_DIR / "_temp_upload.csv"
                temp_path.write_bytes(uploaded_file.getvalue())
                selected_path = str(temp_path)

    if "selected_path" not in locals() or not selected_path:
        st.info("Bitte CSV in der Sidebar auswählen oder hochladen.")
        st.stop()

    st.sidebar.header("Ansicht")
    if "active_prefix" not in st.session_state:
        st.session_state["active_prefix"] = "view_a"
    active_prefix = st.sidebar.selectbox(
        "Filter bearbeiten für",
        options=["view_a", "view_b", "view_c"],
        format_func=lambda x: {"view_a": "Filteransicht A", "view_b": "Filteransicht B", "view_c": "Filteransicht C"}[x],
        key="active_prefix",
    )

    df = load_csv_from_path(selected_path)

    # Apply pending presets before creating filter widgets to avoid session_state conflicts.
    for pfx in ["view_a", "view_b", "view_c"]:
        pending_key = f"{pfx}_pending_preset_state"
        if pending_key in st.session_state:
            apply_view_state(st.session_state.pop(pending_key))

    st.subheader("Spalten der aktuellen CSV")
    st.write(list(df.columns))

    tab1, tab2, tab3, tab4 = st.tabs(["Filteransicht A", "Filteransicht B", "Filteransicht C", "Standort-Analyse"])
    with tab1:
        render_view(df, "Filteransicht A", "view_a", active_prefix)
    with tab2:
        render_view(df, "Filteransicht B", "view_b", active_prefix)
    with tab3:
        render_view(df, "Filteransicht C", "view_c", active_prefix)
    with tab4:
        render_standort_analyse(df)


if __name__ == "__main__":
    main()
