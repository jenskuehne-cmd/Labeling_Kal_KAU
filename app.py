import json
import re
from io import BytesIO
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List
import zipfile

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="KAU Relabeling Priorisierung", layout="wide")

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
REF_DIR = BASE_DIR / "reference_lists"
REF_DIR.mkdir(parents=True, exist_ok=True)
CRITERIA_DIR = BASE_DIR / "criteria_sets"
CRITERIA_DIR.mkdir(parents=True, exist_ok=True)
RUNS_DIR = BASE_DIR / "runs"
RUNS_DIR.mkdir(parents=True, exist_ok=True)


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
        /* Narrow, denser table typography */
        .stDataFrame, .stDataEditor {
            font-family: "Arial Narrow", "Aptos Narrow", "Liberation Sans Narrow", "Noto Sans", sans-serif !important;
            font-size: 12px !important;
        }
        .stDataFrame [role="columnheader"],
        .stDataFrame [role="gridcell"],
        .stDataEditor [role="columnheader"],
        .stDataEditor [role="gridcell"] {
            font-family: "Arial Narrow", "Aptos Narrow", "Liberation Sans Narrow", "Noto Sans", sans-serif !important;
            font-size: 12px !important;
            line-height: 1.15 !important;
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


def build_column_config(df: pd.DataFrame, allow_manual_edit: bool) -> Dict[str, Any]:
    config: Dict[str, Any] = {}
    for col in df.columns:
        label = short_label(col)
        if col == "prio_manual":
            config[col] = st.column_config.SelectboxColumn(
                label=label,
                options=["", "Low", "Medium", "High"],
                help=col,
            )
        elif col == "prio_note":
            config[col] = st.column_config.TextColumn(label=label, help=col)
        elif pd.api.types.is_bool_dtype(df[col]):
            config[col] = st.column_config.CheckboxColumn(label=label, help=col)
        elif pd.api.types.is_numeric_dtype(df[col]):
            config[col] = st.column_config.NumberColumn(label=label, help=col)
        else:
            config[col] = st.column_config.TextColumn(label=label, help=col)
    return config


def empty_value_mask(series: pd.Series) -> pd.Series:
    s = series.astype("string")
    normalized = s.str.strip().str.lower()
    placeholders = {"", "none", "null", "nan", "na", "n/a"}
    return series.isna() | normalized.isin(placeholders)


def json_safe(value: Any) -> Any:
    if isinstance(value, (date, datetime, pd.Timestamp)):
        return value.isoformat()
    if isinstance(value, pd.DataFrame):
        return value.to_dict(orient="records")
    if isinstance(value, pd.Series):
        return value.tolist()
    if isinstance(value, tuple):
        return [json_safe(v) for v in value]
    if isinstance(value, list):
        return [json_safe(v) for v in value]
    if isinstance(value, dict):
        return {k: json_safe(v) for k, v in value.items()}
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    return value


def is_forbidden_widget_key(key: str) -> bool:
    return (
        "_preset_" in key
        or "_editor" in key
        or "__filter_store" in key
        or key.endswith("_qc_excluded_df")
        or key.endswith("_criteria_excluded_df")
        or key.endswith("_excluded_df")
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


@st.cache_data(show_spinner=False)
def load_qc_green_assets_from_xlsx(path_str: str, color_hex: str, asset_col_name: str) -> List[str]:
    wb = load_workbook(path_str, data_only=True)
    ws = wb.active
    headers = [ws.cell(1, c).value for c in range(1, ws.max_column + 1)]
    normalized = [str(h).strip() if h is not None else "" for h in headers]
    target = str(asset_col_name).strip()
    if target not in normalized:
        return []
    asset_idx = normalized.index(target) + 1

    green = color_hex.upper().replace("#", "")
    assets: List[str] = []
    for r in range(2, ws.max_row + 1):
        row_has_green = False
        for c in range(1, ws.max_column + 1):
            cell = ws.cell(r, c)
            rgb = cell.fill.fgColor.rgb
            if isinstance(rgb, str) and rgb.upper().endswith(green):
                row_has_green = True
                break
        if row_has_green:
            value = ws.cell(r, asset_idx).value
            if value is not None and str(value).strip():
                assets.append(str(value).strip())
    return sorted(set(assets))


def list_criteria_sets() -> List[Path]:
    return sorted(CRITERIA_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)


def default_criteria_set() -> Dict[str, Any]:
    return {
        "name": "baseline_v1",
        "version": datetime.now().strftime("%Y-%m-%d_%H%M%S"),
        "notes": "Initiales Kriterien-Set",
        "rules": [
            {
                "id": "interval_gt_10",
                "active": True,
                "type": "numeric_gt",
                "column": "Interval",
                "value": 10,
                "action": "exclude_from_prio",
                "reason": "Intervall > 10 Monate",
            },
            {
                "id": "asset_id_len_lt_34",
                "active": True,
                "type": "string_length_lt",
                "column": "Asset ID",
                "value": 34,
                "action": "exclude_from_prio",
                "reason": "Asset ID unter 34 Zeichen",
            },
            {
                "id": "qc_self_labeled",
                "active": True,
                "type": "ref_list_match",
                "column": "Asset ID",
                "ref_list": "qc_self_labeled_assets",
                "action": "exclude_from_prio",
                "reason": "Wird durch QC selbst gelabelt",
            },
            {
                "id": "prio1_due_soon",
                "active": True,
                "action": "assign_prio",
                "priority": "P1",
                "reason": "Due Date ist kurzfristig",
                "when": {"op": "<=", "column": "months_until_due", "value": 1.0},
            },
            {
                "id": "prio2_within_interval_half",
                "active": True,
                "action": "assign_prio",
                "priority": "P2",
                "reason": "Due Date liegt innerhalb Intervall/2",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": ">", "column": "months_until_due", "value": 1.0},
                        {"op": "<=", "column": "months_until_due", "value_col": "interval_half_months"},
                    ],
                },
            },
        ],
    }


def ensure_default_criteria_file() -> Path:
    existing = list_criteria_sets()
    if existing:
        return existing[0]
    p = CRITERIA_DIR / "baseline_v1.json"
    p.write_text(json.dumps(default_criteria_set(), ensure_ascii=False, indent=2), encoding="utf-8")
    return p


def load_criteria(path_str: str) -> Dict[str, Any]:
    try:
        return json.loads(Path(path_str).read_text(encoding="utf-8"))
    except Exception:
        return default_criteria_set()


def save_criteria(criteria: Dict[str, Any], filename: str) -> Path:
    safe = filename.strip() or f"criteria_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    if not safe.endswith(".json"):
        safe += ".json"
    target = CRITERIA_DIR / safe
    target.write_text(json.dumps(criteria, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def add_derived_time_columns(df: pd.DataFrame) -> pd.DataFrame:
    working = df.copy()

    due_candidates = ["Due Date BMRAM / End Datum SAP", "Due Date", "End Datum SAP"]
    start_candidates = ["Start Datum SAP Auftrag", "Start Datum", "GoLive", "Go Live"]
    interval_candidates = ["Interval", "Intervall"]

    due_col = next((c for c in due_candidates if c in working.columns), None)
    start_col = next((c for c in start_candidates if c in working.columns), None)
    interval_col = next((c for c in interval_candidates if c in working.columns), None)

    today = pd.Timestamp.today().normalize()

    if due_col:
        due_dt = pd.to_datetime(working[due_col], errors="coerce", dayfirst=True)
        working["months_until_due"] = (due_dt - today).dt.days / 30.4375
    else:
        working["months_until_due"] = pd.NA

    if start_col:
        start_dt = pd.to_datetime(working[start_col], errors="coerce", dayfirst=True)
        working["months_since_start"] = (today - start_dt).dt.days / 30.4375
    else:
        working["months_since_start"] = pd.NA

    if interval_col:
        interval_num = pd.to_numeric(working[interval_col], errors="coerce")
        working["interval_half_months"] = interval_num / 2.0
    else:
        working["interval_half_months"] = pd.NA

    return working


def _compare_series(left: pd.Series, op: str, right: Any) -> pd.Series:
    if isinstance(right, pd.Series):
        right_series = right
    else:
        right_series = pd.Series([right] * len(left), index=left.index)

    if op == ">":
        return left > right_series
    if op == ">=":
        return left >= right_series
    if op == "<":
        return left < right_series
    if op == "<=":
        return left <= right_series
    if op == "==":
        return left == right_series
    if op == "!=":
        return left != right_series
    return pd.Series([False] * len(left), index=left.index)


def evaluate_condition(df: pd.DataFrame, cond: Dict[str, Any], reference_lists: Dict[str, set]) -> pd.Series:
    if not isinstance(cond, dict):
        return pd.Series([False] * len(df), index=df.index)

    op = str(cond.get("op", "")).strip().upper()

    if op == "AND":
        conditions = cond.get("conditions", []) or []
        if not conditions:
            return pd.Series([True] * len(df), index=df.index)
        mask = pd.Series([True] * len(df), index=df.index)
        for sub in conditions:
            mask = mask & evaluate_condition(df, sub, reference_lists).fillna(False)
        return mask

    if op == "OR":
        conditions = cond.get("conditions", []) or []
        if not conditions:
            return pd.Series([False] * len(df), index=df.index)
        mask = pd.Series([False] * len(df), index=df.index)
        for sub in conditions:
            mask = mask | evaluate_condition(df, sub, reference_lists).fillna(False)
        return mask

    if op == "NOT":
        sub = cond.get("condition", {})
        return ~evaluate_condition(df, sub, reference_lists).fillna(False)

    col = cond.get("column")
    if col not in df.columns:
        return pd.Series([False] * len(df), index=df.index)

    left_num = pd.to_numeric(df[col], errors="coerce")
    right_col = cond.get("value_col")
    if right_col and right_col in df.columns:
        right_num = pd.to_numeric(df[right_col], errors="coerce")
        return _compare_series(left_num, op, right_num).fillna(False)

    value = cond.get("value")
    if op in {">", ">=", "<", "<="}:
        try:
            return _compare_series(left_num, op, float(value)).fillna(False)
        except Exception:
            return pd.Series([False] * len(df), index=df.index)

    if op in {"==", "!="}:
        if isinstance(value, (int, float)):
            return _compare_series(left_num, op, value).fillna(False)
        left_text = df[col].astype("string")
        right_text = "" if value is None else str(value)
        return _compare_series(left_text, op, right_text).fillna(False)

    if op == "REGEX":
        pattern = str(cond.get("pattern", "")).strip()
        if not pattern:
            return pd.Series([False] * len(df), index=df.index)
        try:
            re.compile(pattern)
            return df[col].astype("string").str.contains(pattern, case=False, regex=True, na=False)
        except re.error:
            return pd.Series([False] * len(df), index=df.index)

    if op == "IN_REF_LIST":
        ref_name = str(cond.get("ref_list", "")).strip()
        ref_values = reference_lists.get(ref_name, set())
        if not ref_values:
            return pd.Series([False] * len(df), index=df.index)
        return df[col].astype("string").str.strip().isin(ref_values)

    return pd.Series([False] * len(df), index=df.index)


def apply_criteria_rules(df: pd.DataFrame, criteria: Dict[str, Any], reference_lists: Dict[str, set]) -> tuple[pd.DataFrame, pd.DataFrame]:
    if not criteria or not criteria.get("rules"):
        included = df.copy()
        included["decision_status"] = "in_prio"
        included["exclude_reasons"] = ""
        included["matched_rule_ids"] = ""
        return included, df.iloc[0:0].copy()

    working = add_derived_time_columns(df)
    reason_col = pd.Series([""] * len(working), index=working.index, dtype="string")
    rule_col = pd.Series([""] * len(working), index=working.index, dtype="string")
    exclude_mask = pd.Series([False] * len(working), index=working.index)
    prio_stage = pd.Series(["P3"] * len(working), index=working.index, dtype="string")
    prio_rank = pd.Series([3] * len(working), index=working.index, dtype="int64")

    for rule in criteria.get("rules", []):
        if not rule.get("active", True):
            continue
        col = rule.get("column")
        rtype = str(rule.get("type", "")).strip()
        reason = str(rule.get("reason", "")).strip() or str(rule.get("id", "rule"))
        rid = str(rule.get("id", "rule"))
        action = str(rule.get("action", "exclude_from_prio"))
        mask = pd.Series([False] * len(working), index=working.index)

        if isinstance(rule.get("when"), dict):
            mask = evaluate_condition(working, rule.get("when"), reference_lists).fillna(False)
        else:
            if col not in working.columns:
                continue
            # Legacy flat rules remain supported.
            if rtype == "numeric_gt":
                val = float(rule.get("value", 0))
                mask = pd.to_numeric(working[col], errors="coerce") > val
            elif rtype == "string_length_lt":
                val = int(rule.get("value", 0))
                mask = working[col].astype("string").str.len().fillna(0) < val
            elif rtype == "regex_match":
                pattern = str(rule.get("pattern", "")).strip()
                if pattern:
                    try:
                        re.compile(pattern)
                        mask = working[col].astype("string").str.contains(pattern, case=False, regex=True, na=False)
                    except re.error:
                        pass
            elif rtype == "date_between":
                start = str(rule.get("start", "")).strip()
                end = str(rule.get("end", "")).strip()
                parsed = pd.to_datetime(working[col], errors="coerce")
                if start and end:
                    try:
                        d1 = pd.to_datetime(start)
                        d2 = pd.to_datetime(end)
                        mask = parsed.between(d1, d2)
                    except Exception:
                        pass
            elif rtype == "ref_list_match":
                ref_name = str(rule.get("ref_list", "")).strip()
                ref_values = reference_lists.get(ref_name, set())
                if ref_values:
                    mask = working[col].astype("string").str.strip().isin(ref_values)

        if action == "exclude_from_prio":
            exclude_mask = exclude_mask | mask.fillna(False)
            reason_col.loc[mask] = (reason_col.loc[mask] + "; " + reason).str.strip("; ")
            rule_col.loc[mask] = (rule_col.loc[mask] + "; " + rid).str.strip("; ")
        elif action == "assign_prio":
            target = str(rule.get("priority", "P3")).upper()
            target_rank = {"P1": 1, "P2": 2, "P3": 3}.get(target, 3)
            upgrade = mask.fillna(False) & (target_rank < prio_rank)
            prio_rank.loc[upgrade] = target_rank
            prio_stage.loc[upgrade] = target
            rule_col.loc[upgrade] = (rule_col.loc[upgrade] + "; " + rid).str.strip("; ")

    excluded = working[exclude_mask].copy()
    excluded["decision_status"] = "excluded"
    excluded["exclude_reasons"] = reason_col.loc[exclude_mask].fillna("")
    excluded["matched_rule_ids"] = rule_col.loc[exclude_mask].fillna("")
    excluded["prio_stage"] = "excluded"

    included = working[~exclude_mask].copy()
    included["decision_status"] = "in_prio"
    included["exclude_reasons"] = ""
    included["matched_rule_ids"] = rule_col.loc[~exclude_mask].fillna("")
    included["prio_stage"] = prio_stage.loc[~exclude_mask].fillna("P3")
    return included, excluded


def _find_rule(criteria: Dict[str, Any], rule_id: str) -> Dict[str, Any] | None:
    for rule in criteria.get("rules", []):
        if str(rule.get("id", "")) == rule_id:
            return rule
    return None


def _count_rule_matches(df: pd.DataFrame, rule: Dict[str, Any], reference_lists: Dict[str, set]) -> int:
    working = add_derived_time_columns(df)
    col = rule.get("column")
    rtype = str(rule.get("type", "")).strip()
    mask = pd.Series([False] * len(working), index=working.index)
    if isinstance(rule.get("when"), dict):
        mask = evaluate_condition(working, rule.get("when"), reference_lists).fillna(False)
        return int(mask.sum())
    if col not in working.columns:
        return 0
    if rtype == "numeric_gt":
        mask = pd.to_numeric(working[col], errors="coerce") > float(rule.get("value", 0))
    elif rtype == "string_length_lt":
        mask = working[col].astype("string").str.len().fillna(0) < int(rule.get("value", 0))
    elif rtype == "regex_match":
        pattern = str(rule.get("pattern", "")).strip()
        if pattern:
            try:
                re.compile(pattern)
                mask = working[col].astype("string").str.contains(pattern, case=False, regex=True, na=False)
            except re.error:
                pass
    elif rtype == "date_between":
        start = str(rule.get("start", "")).strip()
        end = str(rule.get("end", "")).strip()
        if start and end:
            try:
                parsed = pd.to_datetime(working[col], errors="coerce")
                mask = parsed.between(pd.to_datetime(start), pd.to_datetime(end))
            except Exception:
                pass
    elif rtype == "ref_list_match":
        ref_name = str(rule.get("ref_list", "")).strip()
        ref_values = reference_lists.get(ref_name, set())
        if ref_values:
            mask = working[col].astype("string").str.strip().isin(ref_values)
    return int(mask.fillna(False).sum())


def render_live_criteria_sidebar(df: pd.DataFrame) -> None:
    criteria = st.session_state.get("active_criteria_set", default_criteria_set())
    if not criteria.get("rules"):
        return

    ref_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}
    with st.sidebar.expander("Live-Kriterien", expanded=True):
        st.caption(f"Aktiv: {criteria.get('name', '')} | {criteria.get('version', '')}")

        r_interval = _find_rule(criteria, "interval_gt_10")
        if r_interval:
            r_interval["active"] = st.checkbox("Intervall-Regel aktiv", value=bool(r_interval.get("active", True)), key="live_rule_interval_on")
            r_interval["value"] = st.slider("Intervall > (Monate)", min_value=0, max_value=36, value=int(float(r_interval.get("value", 10))), step=1, key="live_rule_interval_val")
            r_interval["reason"] = st.text_input("Grund Intervall-Regel", value=str(r_interval.get("reason", "")), key="live_rule_interval_reason")
            st.caption(f"Treffer aktuell: {_count_rule_matches(df, r_interval, ref_lists)}")

        r_asset = _find_rule(criteria, "asset_id_len_lt_34")
        if r_asset:
            r_asset["active"] = st.checkbox("Asset-ID-Längenregel aktiv", value=bool(r_asset.get("active", True)), key="live_rule_asset_on")
            r_asset["value"] = st.slider("Asset ID Länge < ", min_value=1, max_value=80, value=int(float(r_asset.get("value", 34))), step=1, key="live_rule_asset_val")
            r_asset["reason"] = st.text_input("Grund Asset-ID-Regel", value=str(r_asset.get("reason", "")), key="live_rule_asset_reason")
            st.caption(f"Treffer aktuell: {_count_rule_matches(df, r_asset, ref_lists)}")

        r_qc = _find_rule(criteria, "qc_self_labeled")
        if r_qc:
            r_qc["active"] = st.checkbox("QC-Selbstlabeling-Regel aktiv", value=bool(r_qc.get("active", True)), key="live_rule_qc_on")
            r_qc["reason"] = st.text_input("Grund QC-Regel", value=str(r_qc.get("reason", "")), key="live_rule_qc_reason")
            st.caption(f"Treffer aktuell: {_count_rule_matches(df, r_qc, ref_lists)}")

        st.session_state["active_criteria_set"] = criteria
def apply_qc_exclusion(df: pd.DataFrame, prefix: str, value_getter) -> tuple[pd.DataFrame, pd.DataFrame]:
    enabled = bool(value_getter(f"{prefix}_qc_exclude_on", False))
    asset_col = value_getter(f"{prefix}_qc_asset_col", "Asset ID")
    reason = value_getter(f"{prefix}_qc_reason", "Wird durch QC selbst gelabelt")
    qc_assets = st.session_state.get("qc_green_assets", [])
    if (not enabled) or (not qc_assets) or (asset_col not in df.columns):
        return df, df.iloc[0:0].copy()

    asset_series = df[asset_col].astype("string").str.strip()
    mask = asset_series.isin(set(qc_assets))
    excluded = df[mask].copy()
    excluded["exclude_rule"] = "qc_self_labeled"
    excluded["exclude_reason"] = reason
    included = df[~mask].copy()
    return included, excluded


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


def _flatten_loaded_state(raw: Dict[str, Any]) -> Dict[str, Any]:
    flat: Dict[str, Any] = {}

    def walk(obj: Any) -> None:
        if isinstance(obj, dict):
            for k, v in obj.items():
                if k.startswith(("view_a_", "view_b_", "view_c_")) and not is_forbidden_widget_key(k):
                    flat[k] = v
                walk(v)

    walk(raw)
    return flat


def extract_loaded_state(parsed: Dict[str, Any]) -> Dict[str, Any]:
    # Support both "preset" format and "csv+filter export" format.
    candidate = parsed.get("state")
    if not isinstance(candidate, dict):
        candidate = parsed.get("filters")
    if not isinstance(candidate, dict):
        candidate = parsed if isinstance(parsed, dict) else {}
    return _flatten_loaded_state(candidate)


def _resolve_column_name(raw_name: str, available_cols: List[str]) -> str:
    if raw_name in available_cols:
        return raw_name
    raw_trim = str(raw_name).strip()
    if raw_trim in available_cols:
        return raw_trim
    # Fallback: match by trimmed representation.
    for col in available_cols:
        if str(col).strip() == raw_trim:
            return col
    return raw_name


def normalize_loaded_state_columns(state: Dict[str, Any], available_cols: List[str]) -> Dict[str, Any]:
    normalized: Dict[str, Any] = {}
    for key, value in state.items():
        new_key = key
        new_value = value

        # Direct column selectors.
        if key.endswith("_regex_col") or key.endswith("_date_col") or "_rule_col_" in key:
            if isinstance(value, str):
                new_value = _resolve_column_name(value, available_cols)

        # Category keys encode the column name in the key itself.
        if "_cat_" in key:
            prefix, raw_col = key.split("_cat_", 1)
            mapped_col = _resolve_column_name(raw_col, available_cols)
            new_key = f"{prefix}_cat_{mapped_col}"

        normalized[new_key] = new_value

    return normalized


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
            if isinstance(value, pd.DataFrame):
                continue
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

    filtered, excluded_qc = apply_qc_exclusion(filtered, prefix, sget)
    st.session_state[f"{prefix}_qc_excluded_df"] = excluded_qc

    criteria = st.session_state.get("active_criteria_set", {})
    reference_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}
    filtered, excluded_rules = apply_criteria_rules(filtered, criteria, reference_lists)
    st.session_state[f"{prefix}_criteria_excluded_df"] = excluded_rules

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
                incoming_state = extract_loaded_state(parsed)
                incoming_state = remap_state_prefix(incoming_state, prefix)
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

    with st.sidebar.expander("QC-Ausnahme", expanded=False) as qc_box:
        qc_box.checkbox("QC-Selbstlabeling ausschließen", value=False, key=f"{prefix}_qc_exclude_on")
        asset_options = list(filtered.columns)
        default_asset = "Asset ID" if "Asset ID" in asset_options else (asset_options[0] if asset_options else "")
        if asset_options:
            qc_box.selectbox("Asset-Spalte", asset_options, index=asset_options.index(default_asset), key=f"{prefix}_qc_asset_col")
        qc_box.text_input("Ausschluss-Begründung", value="Wird durch QC selbst gelabelt", key=f"{prefix}_qc_reason")
        if st.session_state.get("qc_green_assets"):
            qc_box.caption(f"QC-Referenzliste aktiv: {len(st.session_state['qc_green_assets'])} Asset IDs")
        else:
            qc_box.caption("Keine QC-Referenzliste geladen.")

    filtered, excluded_qc = apply_qc_exclusion(filtered, prefix, st.session_state.get)
    st.session_state[f"{prefix}_qc_excluded_df"] = excluded_qc

    criteria = st.session_state.get("active_criteria_set", {})
    reference_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}
    filtered, excluded_rules = apply_criteria_rules(filtered, criteria, reference_lists)
    st.session_state[f"{prefix}_criteria_excluded_df"] = excluded_rules

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

    st.markdown("#### Gruppendetails")
    group_values = summary[group_col].astype("string").fillna("").tolist()
    if not group_values:
        return
    default_group = group_values[0]
    selected_group = st.selectbox(
        "Gruppe auswählen",
        options=group_values,
        index=0,
        key=f"{prefix}_group_detail_pick",
    )
    detail_df = df[df[group_col].astype("string").fillna("") == selected_group].copy()
    st.write(f"{len(detail_df)} Detailzeilen für Gruppe: {selected_group}")
    default_cols = [c for c in [group_col, "Zugänglichkeit", "Standort"] if c in detail_df.columns]
    shown_cols = st.multiselect(
        "Spalten in Details",
        options=list(detail_df.columns),
        default=default_cols if default_cols else list(detail_df.columns)[: min(8, len(detail_df.columns))],
        key=f"{prefix}_group_detail_cols",
    )
    detail_view = detail_df[shown_cols] if shown_cols else detail_df
    st.dataframe(
        detail_view,
        use_container_width=True,
        height=320,
        column_config=build_column_config(detail_view, allow_manual_edit=False),
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

    if "prio_stage" in edited.columns:
        st.markdown("#### Prioritätsstufen (aus Kriterien-Set)")
        prio_summary = edited.groupby("prio_stage", dropna=False).size().reset_index(name="anzahl")
        prio_summary = prio_summary.sort_values("prio_stage")
        st.dataframe(
            prio_summary,
            use_container_width=True,
            height=180,
            column_config=build_column_config(prio_summary, allow_manual_edit=False),
        )

    excluded_qc = st.session_state.get(f"{prefix}_qc_excluded_df")
    if isinstance(excluded_qc, pd.DataFrame) and not excluded_qc.empty:
        st.markdown("#### Ausgeschlossen durch QC-Regel")
        st.write(f"{len(excluded_qc)} Zeilen wurden bewusst ausgeschlossen.")
        show_cols_default = [c for c in ["Asset ID", "Gebäude / MU", "Standort", "exclude_rule", "exclude_reason"] if c in excluded_qc.columns]
        show_cols = st.multiselect(
            "Spalten in Ausschlussliste",
            options=list(excluded_qc.columns),
            default=show_cols_default if show_cols_default else list(excluded_qc.columns)[: min(8, len(excluded_qc.columns))],
            key=f"{prefix}_qc_excluded_cols",
        )
        ex_view = excluded_qc[show_cols] if show_cols else excluded_qc
        st.dataframe(
            ex_view,
            use_container_width=True,
            height=260,
            column_config=build_column_config(ex_view, allow_manual_edit=False),
        )

    excluded_criteria = st.session_state.get(f"{prefix}_criteria_excluded_df")
    if isinstance(excluded_criteria, pd.DataFrame) and not excluded_criteria.empty:
        st.markdown("#### Ausgeschlossen durch Kriterien-Set")
        st.write(f"{len(excluded_criteria)} Zeilen wurden durch aktive Kriterien ausgeschlossen.")
        default_cols = [c for c in ["Asset ID", "Gebäude / MU", "Standort", "matched_rule_ids", "exclude_reasons"] if c in excluded_criteria.columns]
        sel_cols = st.multiselect(
            "Spalten in Kriterien-Ausschlussliste",
            options=list(excluded_criteria.columns),
            default=default_cols if default_cols else list(excluded_criteria.columns)[: min(8, len(excluded_criteria.columns))],
            key=f"{prefix}_criteria_excluded_cols",
        )
        crit_view = excluded_criteria[sel_cols] if sel_cols else excluded_criteria
        st.dataframe(
            crit_view,
            use_container_width=True,
            height=260,
            column_config=build_column_config(crit_view, allow_manual_edit=False),
        )

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




def build_condition_preview(cond: Dict[str, Any]) -> str:
    if not isinstance(cond, dict):
        return "(leer)"
    op = str(cond.get("op", "")).upper()
    if op in {"AND", "OR"}:
        parts = [build_condition_preview(c) for c in cond.get("conditions", [])]
        return f" ({f' {op} '.join(parts)}) "
    if op == "NOT":
        return f"NOT ({build_condition_preview(cond.get('condition', {}))})"
    col = str(cond.get("column", "?"))
    rhs = f"[{cond.get('value_col')}]" if cond.get("value_col") else repr(cond.get("value"))
    return f"{col} {op} {rhs}"


def render_rule_builder(df: pd.DataFrame) -> None:
    st.markdown("#### Regel-Builder (visuell)")
    st.caption("Neu: beliebig viele Bedingungen je Regel (AND/OR), ohne JSON-Handarbeit.")

    columns = list(df.columns) + ["months_until_due", "months_since_start", "interval_half_months"]
    if not columns:
        st.info("Keine Spalten verfügbar.")
        return

    c1, c2, c3 = st.columns([1, 1, 1])
    action = c1.selectbox("Aktion", ["exclude_from_prio", "assign_prio"], key="rb_action")
    priority = c2.selectbox("Prio (bei assign)", ["P1", "P2", "P3"], key="rb_priority")
    logic = c3.selectbox("Verknüpfung", ["AND", "OR"], key="rb_logic")

    count = st.number_input("Anzahl Bedingungen", min_value=1, max_value=8, value=3, step=1, key="rb_cond_count")
    conditions = []
    operators = [">", ">=", "<", "<=", "==", "!="]

    for i in range(int(count)):
        with st.expander(f"Bedingung {i+1}", expanded=(i < 2)):
            a1, a2, a3, a4 = st.columns([2, 1, 1, 2])
            col = a1.selectbox("Spalte", columns, key=f"rb_col_{i}")
            op = a2.selectbox("Operator", operators, key=f"rb_op_{i}")
            mode = a3.selectbox("Werttyp", ["Konstante", "Spalte"], key=f"rb_mode_{i}")
            if mode == "Konstante":
                raw_val = a4.text_input("Wert", value="0", key=f"rb_val_{i}")
                try:
                    val = float(raw_val)
                except Exception:
                    val = raw_val
                cond = {"op": op, "column": col, "value": val}
            else:
                val_col = a4.selectbox("Vergleichsspalte", columns, key=f"rb_val_col_{i}")
                cond = {"op": op, "column": col, "value_col": val_col}
            conditions.append(cond)

    rule_id = st.text_input("Regel-ID", value=f"rule_{datetime.now().strftime('%H%M%S')}", key="rb_id")
    reason = st.text_input("Begründung", value="Manuell erstellte Regel", key="rb_reason")

    when = {"op": logic, "conditions": conditions}
    st.code(json.dumps(when, ensure_ascii=False, indent=2), language="json")
    st.caption(f"Vorschau: {build_condition_preview(when)}")

    if st.button("Regel zum aktiven Kriterien-Set hinzufügen", key="rb_add"):
        criteria = st.session_state.get("active_criteria_set", default_criteria_set())
        new_rule = {
            "id": rule_id.strip() or f"rule_{datetime.now().strftime('%Y%m%d_%H%M%S')}",
            "active": True,
            "action": action,
            "reason": reason,
            "when": when,
        }
        if action == "assign_prio":
            new_rule["priority"] = priority
        criteria.setdefault("rules", []).append(new_rule)
        st.session_state["active_criteria_set"] = criteria
        st.success("Regel hinzugefügt. Jetzt unten 'Kriterien-Version speichern' klicken.")
        st.rerun()


def render_criteria_editor(df: pd.DataFrame) -> None:
    st.markdown("### Kriterien-Set")
    render_rule_builder(df)
    criteria = st.session_state.get("active_criteria_set", default_criteria_set())
    st.write(f"Aktiv: {criteria.get('name', '')} | Version: {criteria.get('version', '')}")
    st.text_area("Notizen", value=str(criteria.get("notes", "")), key="criteria_notes")

    rules = criteria.get("rules", [])
    rules_df = pd.DataFrame(rules if rules else [default_criteria_set()["rules"][0]])
    edited_rules = st.data_editor(
        rules_df,
        use_container_width=True,
        num_rows="dynamic",
        key="criteria_rules_editor",
    )

    c1, c2 = st.columns([2, 1])
    filename = c1.text_input("Dateiname für neue Version", value=f"{criteria.get('name', 'criteria')}_{datetime.now().strftime('%Y%m%d_%H%M')}.json", key="criteria_save_name")
    if c2.button("Kriterien-Version speichern", key="criteria_save_btn"):
        new_set = {
            "name": criteria.get("name", "criteria_set"),
            "version": datetime.now().strftime("%Y-%m-%d_%H%M%S"),
            "notes": st.session_state.get("criteria_notes", ""),
            "rules": edited_rules.fillna("").to_dict(orient="records"),
        }
        target = save_criteria(new_set, filename)
        st.session_state["active_criteria_set"] = new_set
        st.success(f"Gespeichert: {target}")

    st.caption("Regeltypen: numeric_gt, string_length_lt, regex_match, date_between, ref_list_match oder when-Block mit AND/OR/NOT")
    st.caption("Action: exclude_from_prio oder assign_prio (priority: P1/P2/P3)")


def main() -> None:
    apply_custom_style()
    sanitize_filter_stores()
    st.title("KAU Relabeling Priorisierung")

    # Initialize default criteria set and persisted QC reference list.
    if "criteria_initialized" not in st.session_state:
        default_path = ensure_default_criteria_file()
        st.session_state["active_criteria_path"] = str(default_path)
        st.session_state["active_criteria_set"] = load_criteria(str(default_path))
        qc_csv = REF_DIR / "qc_self_labeled_assets.csv"
        if qc_csv.exists():
            try:
                qc_assets = pd.read_csv(qc_csv)["Asset ID"].dropna().astype(str).str.strip().tolist()
                st.session_state["qc_green_assets"] = sorted(set(qc_assets))
                st.session_state["qc_green_source"] = str(qc_csv)
            except Exception:
                pass
        st.session_state["criteria_initialized"] = True

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

    st.sidebar.header("Referenzlisten")
    with st.sidebar.expander("QC-Grünliste (Excel)", expanded=False):
        qc_file = st.file_uploader("QC-Excel hochladen", type=["xlsx"], key="qc_excel_upload")
        qc_color = st.text_input("Grün-Farbcode", value="FF92D050", key="qc_green_color")
        qc_asset_col = st.text_input("Asset-ID Spaltenname", value="Asset ID", key="qc_asset_col_name")
        if qc_file is not None and st.button("QC-Liste einlesen", key="qc_load_btn"):
            qc_path = REF_DIR / Path(qc_file.name).name
            qc_path.write_bytes(qc_file.getvalue())
            assets = load_qc_green_assets_from_xlsx(str(qc_path), qc_color, qc_asset_col)
            st.session_state["qc_green_assets"] = assets
            st.session_state["qc_green_source"] = str(qc_path)
            pd.DataFrame({"Asset ID": assets}).to_csv(REF_DIR / "qc_self_labeled_assets.csv", index=False)
            st.success(f"QC-Liste geladen: {len(assets)} Asset IDs")
        if st.session_state.get("qc_green_assets") is not None:
            st.caption(f"Aktive QC-Liste: {len(st.session_state.get('qc_green_assets', []))} Asset IDs")
            if st.session_state.get("qc_green_source"):
                st.caption(f"Quelle: {st.session_state['qc_green_source']}")

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

    st.sidebar.header("Kriterien")
    criteria_files = list_criteria_sets()
    if criteria_files:
        options = {p.name: str(p) for p in criteria_files}
        current_path = st.session_state.get("active_criteria_path", str(criteria_files[0]))
        current_name = Path(current_path).name if current_path else criteria_files[0].name
        if current_name not in options:
            current_name = criteria_files[0].name
        selected_criteria_name = st.sidebar.selectbox("Aktives Kriterien-Set", list(options.keys()), index=list(options.keys()).index(current_name))
        selected_criteria_path = options[selected_criteria_name]
        if selected_criteria_path != st.session_state.get("active_criteria_path"):
            st.session_state["active_criteria_path"] = selected_criteria_path
            st.session_state["active_criteria_set"] = load_criteria(selected_criteria_path)
            st.rerun()

    df = load_csv_from_path(selected_path)
    render_live_criteria_sidebar(df)

    # Apply pending presets before creating filter widgets to avoid session_state conflicts.
    for pfx in ["view_a", "view_b", "view_c"]:
        pending_key = f"{pfx}_pending_preset_state"
        if pending_key in st.session_state:
            raw_state = st.session_state.pop(pending_key)
            mapped_state = normalize_loaded_state_columns(raw_state, list(df.columns))
            apply_view_state(mapped_state)

    st.subheader("Spalten der aktuellen CSV")
    st.write(list(df.columns))

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["Filteransicht A", "Filteransicht B", "Filteransicht C", "Standort-Analyse", "Kriterien"])
    with tab1:
        render_view(df, "Filteransicht A", "view_a", active_prefix)
    with tab2:
        render_view(df, "Filteransicht B", "view_b", active_prefix)
    with tab3:
        render_view(df, "Filteransicht C", "view_c", active_prefix)
    with tab4:
        render_standort_analyse(df)
    with tab5:
        render_criteria_editor(df)


if __name__ == "__main__":
    main()
