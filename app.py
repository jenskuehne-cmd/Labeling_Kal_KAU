import json
import ast
import math
import re
from html import escape as html_escape
from io import BytesIO
from datetime import date, datetime
from pathlib import Path
from typing import Any, Dict, List
import zipfile

import pandas as pd
import numpy as np
import streamlit as st
from openpyxl import Workbook, load_workbook


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
LAST_SESSION_FILE = RUNS_DIR / "last_session_state.json"


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
            font-size: 0.875rem;
        }
        /* Narrow, denser table typography */
        .stDataFrame, .stDataEditor {
            font-family: "Arial Narrow", "Aptos Narrow", "Liberation Sans Narrow", "Noto Sans", sans-serif !important;
            font-size: 10px !important;
        }
        .stDataFrame [role="columnheader"],
        .stDataFrame [role="gridcell"],
        .stDataEditor [role="columnheader"],
        .stDataEditor [role="gridcell"] {
            font-family: "Arial Narrow", "Aptos Narrow", "Liberation Sans Narrow", "Noto Sans", sans-serif !important;
            font-size: 10px !important;
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




def normalize_name_for_match(s: str) -> str:
    t = str(s).strip().lower()
    t = t.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    t = re.sub(r"[^a-z0-9]+", "", t)
    return t


def find_column_by_candidates(df: pd.DataFrame, candidates: List[str]) -> str | None:
    cols = list(df.columns)
    for c in candidates:
        if c in cols:
            return c
    normalized = {normalize_name_for_match(col): col for col in cols}
    for c in candidates:
        key = normalize_name_for_match(c)
        if key in normalized:
            return normalized[key]
    return None

def build_column_config(df: pd.DataFrame, allow_manual_edit: bool) -> Dict[str, Any]:
    config: Dict[str, Any] = {}
    for col in df.columns:
        label = short_label(col)
        col_lower = str(col).lower()
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
            if any(token in col_lower for token in ["personentage", "wochenbedarf"]):
                fmt = "%.1f"
            elif any(token in col_lower for token in ["aufwand", "minuten", "minutes"]):
                fmt = "%.0f"
            elif any(token in col_lower for token in ["puffer", "prozent", "percent", "pct"]):
                fmt = "%.0f"
            elif pd.api.types.is_integer_dtype(df[col]):
                fmt = "%d"
            else:
                fmt = "%.2f"
            config[col] = st.column_config.NumberColumn(label=label, help=col, format=fmt)
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
        or key.endswith("_download_zip")
        or key.endswith("_export")
        or key.endswith("_export_zip")
        or key.endswith("_activate_sidebar")
        or key.endswith("_pending_preset_state")
        or key.endswith("_preset_loaded_msg")
        or key.endswith("_summary_preview")
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


def _session_restore_value(key: str, value: Any) -> Any:
    if key == "golive_reference_date" and isinstance(value, str):
        try:
            return date.fromisoformat(value)
        except Exception:
            return date(2026, 8, 10)
    return value


def restore_last_session() -> None:
    if st.session_state.get("_last_session_loaded"):
        return
    st.session_state["_last_session_loaded"] = True
    if not LAST_SESSION_FILE.exists():
        return
    try:
        payload = json.loads(LAST_SESSION_FILE.read_text(encoding="utf-8"))
    except Exception as exc:
        st.session_state["_last_session_restore_error"] = str(exc)
        return

    for key, value in payload.get("app_state", {}).items():
        if key not in st.session_state and not is_forbidden_widget_key(key):
            st.session_state[key] = _session_restore_value(key, value)

    for prefix, state in payload.get("filter_stores", {}).items():
        if prefix in {"view_a", "view_b", "view_c"} and isinstance(state, dict):
            set_filter_store(prefix, state)

    for key, value in payload.get("standort_state", {}).items():
        if key not in st.session_state and key.startswith("sa_") and not is_forbidden_widget_key(key):
            st.session_state[key] = value

    if payload.get("selected_csv_name") and "selected_csv_name" not in st.session_state:
        st.session_state["selected_csv_name"] = payload.get("selected_csv_name")
    if payload.get("selected_csv_path"):
        st.session_state["last_selected_csv_path"] = payload.get("selected_csv_path")
    if payload.get("decision_tree_config"):
        set_active_decision_tree_config(payload.get("decision_tree_config"))
    st.session_state["_last_session_saved_at"] = payload.get("saved_at", "")


def collect_persisted_session_state(selected_path: str | None) -> Dict[str, Any]:
    app_keys = [
        "sidebar_mode",
        "active_prefix",
        "active_criteria_path",
        "decision_tree_config",
        "golive_reference_date",
        "cap_min_normal",
        "cap_min_medium",
        "cap_buffer_pct",
        "cap_persons",
        "cap_hours_day",
    ]
    app_state = {
        key: json_safe(st.session_state.get(key))
        for key in app_keys
        if key in st.session_state and not is_forbidden_widget_key(key)
    }
    filter_stores = {
        prefix: get_filter_store(prefix)
        for prefix in ["view_a", "view_b", "view_c"]
        if get_filter_store(prefix)
    }
    standort_state = {
        key: json_safe(value)
        for key, value in st.session_state.items()
        if key.startswith("sa_") and not is_forbidden_widget_key(key)
    }
    return {
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "selected_csv_path": selected_path or st.session_state.get("last_selected_csv_path", ""),
        "selected_csv_name": st.session_state.get("selected_csv_name", ""),
        "app_state": app_state,
        "filter_stores": json_safe(filter_stores),
        "standort_state": json_safe(standort_state),
    }


def save_last_session(selected_path: str | None) -> None:
    try:
        LAST_SESSION_FILE.write_text(
            json.dumps(collect_persisted_session_state(selected_path), ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except Exception as exc:
        st.session_state["_last_session_save_error"] = str(exc)


def reset_last_session() -> None:
    try:
        if LAST_SESSION_FILE.exists():
            LAST_SESSION_FILE.unlink()
    except Exception as exc:
        st.session_state["_last_session_reset_error"] = str(exc)


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


def build_named_export_zip(files: Dict[str, bytes | str]) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for name, content in files.items():
            if isinstance(content, str):
                content = content.encode("utf-8")
            zf.writestr(name, content)
    return buf.getvalue()


def _flatten_for_rows(value: Any, prefix: str = "") -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []
    if isinstance(value, dict):
        for key, item in value.items():
            child_prefix = f"{prefix}.{key}" if prefix else str(key)
            rows.extend(_flatten_for_rows(item, child_prefix))
    elif isinstance(value, list):
        for idx, item in enumerate(value):
            child_prefix = f"{prefix}[{idx}]"
            rows.extend(_flatten_for_rows(item, child_prefix))
    else:
        rows.append({"key": prefix, "value": value})
    return rows


def _sanitize_excel_scalar(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, (date, datetime, pd.Timestamp)):
        return value.isoformat()
    if hasattr(value, "item") and not isinstance(value, (str, bytes)):
        try:
            value = value.item()
        except Exception:
            pass
    if isinstance(value, (dict, list, tuple, set)):
        try:
            value = json.dumps(json_safe(value), ensure_ascii=False)
        except Exception:
            value = str(value)
    if isinstance(value, str):
        cleaned = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", " ", value)
        if cleaned[:1] in {"=", "+", "-", "@"}:
            cleaned = "'" + cleaned
        if len(cleaned) > 32760:
            cleaned = cleaned[:32760] + "..."
        return cleaned
    return value


def _sanitize_excel_frame(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df.copy() if isinstance(df, pd.DataFrame) else pd.DataFrame()
    cleaned = df.copy()
    for col in cleaned.columns:
        if cleaned[col].dtype == "object" or pd.api.types.is_string_dtype(cleaned[col]):
            cleaned[col] = cleaned[col].map(_sanitize_excel_scalar)
    return cleaned


def build_phase_plan_excel_export(grp: pd.DataFrame, raw_export: pd.DataFrame, export_meta: Dict[str, Any]) -> bytes:
    buf = BytesIO()
    wb = Workbook(write_only=True)

    def write_sheet(sheet_name: str, frame: pd.DataFrame) -> None:
        ws = wb.create_sheet(title=sheet_name)
        frame = _sanitize_excel_frame(frame)
        ws.append([_sanitize_excel_scalar(col) for col in frame.columns.tolist()])
        for row in frame.itertuples(index=False, name=None):
            ws.append([_sanitize_excel_scalar(v) for v in row])

    def write_kv_sheet(sheet_name: str, rows: List[Dict[str, Any]]) -> None:
        ws = wb.create_sheet(title=sheet_name)
        ws.append(["Schluessel", "Wert"])
        for row in rows:
            ws.append([
                _sanitize_excel_scalar(row.get("key", "")),
                _sanitize_excel_scalar(row.get("value", "")),
            ])

    def write_cover_sheet() -> None:
        ws = wb.create_sheet(title="Start")
        ws.append(["Kennzahl", "Wert"])
        cover_rows = [
            ("Phasenplan-Export", ""),
            ("Erstellt am", export_meta.get("generated_at", "")),
            ("Quelle CSV", export_meta.get("source_csv_name", "")),
            ("Entscheidungsbaum", export_meta.get("decision_tree_name", "")),
            ("Entscheidungsbaum-Version", export_meta.get("decision_tree_version", "")),
            ("Phasen", export_meta.get("summary", {}).get("phases", "")),
            ("Messstellen", export_meta.get("summary", {}).get("rows", "")),
            ("Gesamt-Personentage", export_meta.get("summary", {}).get("total_personentage", "")),
            ("Gesamt-Wochenbedarf", export_meta.get("summary", {}).get("total_wochenbedarf", "")),
        ]
        for key, value in cover_rows:
            ws.append([_sanitize_excel_scalar(key), _sanitize_excel_scalar(value)])

    write_cover_sheet()
    write_sheet("phase_plan_aggregated", grp)
    write_sheet("raw_classified", raw_export)

    meta_rows = _flatten_for_rows(export_meta)
    write_kv_sheet("metadata", meta_rows)

    criteria = st.session_state.get("active_criteria_set", {})
    rules = criteria.get("rules", []) if isinstance(criteria, dict) else []
    rules_df = pd.DataFrame(rules if isinstance(rules, list) else [])
    if not rules_df.empty:
        rename_map = {
            "id": "Regel-ID",
            "active": "Aktiv",
            "action": "Aktion",
            "reason": "Begruendung",
            "priority": "Prioritaet",
            "priority_substage": "Unterstufe",
            "relabel_phase": "Phase",
            "unterbruch_erforderlich": "Unterbruch",
            "recommended_window": "Empfohlenes Fenster",
        }
        rules_df = rules_df.rename(columns={k: v for k, v in rename_map.items() if k in rules_df.columns})
    write_sheet("decision_tree_rules", rules_df)

    wb.save(buf)
    return buf.getvalue()


def summarize_decision_tree(criteria: Dict[str, Any]) -> Dict[str, Any]:
    rules = []
    for rule in criteria.get("rules", []) if isinstance(criteria, dict) else []:
        if not isinstance(rule, dict):
            continue
        rules.append(
            {
                "id": rule.get("id", ""),
                "active": bool(rule.get("active", True)),
                "action": rule.get("action", ""),
                "priority": rule.get("priority", ""),
                "priority_substage": rule.get("priority_substage", rule.get("substage", "")),
                "relabel_phase": rule.get("relabel_phase", rule.get("phase", "")),
                "unterbruch_erforderlich": bool(rule.get("unterbruch_erforderlich", rule.get("mark_shutdown", False))),
                "reason": rule.get("reason", ""),
                "when": rule.get("when", {}),
            }
        )
    return {
        "name": str(criteria.get("name", "")) if isinstance(criteria, dict) else "",
        "version": str(criteria.get("version", "")) if isinstance(criteria, dict) else "",
        "notes": str(criteria.get("notes", "")) if isinstance(criteria, dict) else "",
        "config": json_safe(criteria.get("config", {})) if isinstance(criteria, dict) else {},
        "rule_count": len(rules),
        "active_rule_count": sum(1 for r in rules if r["active"]),
        "rules": rules,
    }


def _phase_sort_key(phase: Any) -> tuple:
    text = str(phase).strip()
    low = text.lower()
    if low == "excluded":
        return (99, low)
    if "shutdown_26" in low:
        return (10, low)
    if "postgl_bis_oktober" in low:
        return (20, low)
    if "klaerung" in low or "zugaenglichkeit" in low:
        return (30, low)
    if "shutdown_27" in low:
        return (40, low)
    if "postgl_bis_januar" in low:
        return (50, low)
    if "opportunistisch" in low:
        return (60, low)
    if "backlog" in low:
        return (70, low)
    return (80, low)


def build_phase_overview_tables(grp: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    overview: Dict[str, pd.DataFrame] = {}
    if grp is None or grp.empty:
        return overview

    prio_order = ["P1", "P2", "P2A", "P3", "P4", "P5", "P6"]
    prio_counts = (
        grp.groupby("prio_stage", dropna=False)[["anzahl_messstellen", "aufwand_plus_puffer_min", "personentage"]]
        .sum()
        .reindex(prio_order, fill_value=0)
        .reset_index()
        .rename(columns={"prio_stage": "prio_stage", "anzahl_messstellen": "messstellen"})
    )
    overview["prio_counts"] = prio_counts

    phase_counts = (
        grp.groupby("relabel_phase", dropna=False)[["anzahl_messstellen", "aufwand_plus_puffer_min", "personentage"]]
        .sum()
        .reset_index()
        .sort_values("relabel_phase", key=lambda s: s.map(_phase_sort_key))
        .rename(columns={"relabel_phase": "phase", "anzahl_messstellen": "messstellen"})
    )
    overview["phase_counts"] = phase_counts

    phase_prio = (
        grp.pivot_table(index="relabel_phase", columns="prio_stage", values="anzahl_messstellen", aggfunc="sum", fill_value=0)
        .reindex(index=sorted(grp["relabel_phase"].dropna().unique(), key=_phase_sort_key))
        .reindex(columns=[p for p in prio_order if p in grp["prio_stage"].dropna().unique()], fill_value=0)
        .reset_index()
        .rename(columns={"relabel_phase": "phase"})
    )
    overview["phase_prio_matrix"] = phase_prio

    return overview


def _nice_tick_step(max_value: float, target_ticks: int = 4) -> int:
    if not max_value or max_value <= 0:
        return 1
    raw_step = max_value / max(target_ticks, 1)
    magnitude = 10 ** max(int(math.floor(math.log10(raw_step))), 0)
    residual = raw_step / magnitude
    if residual <= 1:
        step = 1 * magnitude
    elif residual <= 2:
        step = 2 * magnitude
    elif residual <= 5:
        step = 5 * magnitude
    else:
        step = 10 * magnitude
    return max(int(step), 1)


def _build_overview_svg(
    rows: pd.DataFrame,
    label_col: str,
    count_col: str,
    effort_col: str,
    mode: str,
    title: str,
    effort_label: str = "Aufwand (h)",
) -> str:
    if rows is None or rows.empty:
        return ""

    data = rows[[label_col, count_col, effort_col]].copy()
    data[count_col] = pd.to_numeric(data[count_col], errors="coerce").fillna(0)
    data[effort_col] = pd.to_numeric(data[effort_col], errors="coerce").fillna(0)
    labels = [str(v) for v in data[label_col].tolist()]
    counts = [float(v) for v in data[count_col].tolist()]
    efforts = [float(v) / 60.0 for v in data[effort_col].tolist()]

    count_max = max(max(counts), 1.0)
    effort_max = max(max(efforts), 1.0)
    label_count = len(labels)
    slot_width = 92
    width = max(760, 120 + label_count * slot_width)
    height = 360
    margin_left = 68
    margin_right = 72
    margin_top = 44
    margin_bottom = 74
    plot_width = width - margin_left - margin_right
    plot_height = height - margin_top - margin_bottom
    base_y = margin_top + plot_height

    left_step = _nice_tick_step(count_max)
    left_ticks = list(range(0, int(math.ceil(count_max / left_step) * left_step) + left_step, left_step))
    right_step = _nice_tick_step(effort_max)
    right_ticks = list(range(0, int(math.ceil(effort_max / right_step) * right_step) + right_step, right_step))
    if len(left_ticks) > 6:
        left_ticks = left_ticks[::2]
    if len(right_ticks) > 6:
        right_ticks = right_ticks[::2]

    def y_left(value: float) -> float:
        return margin_top + plot_height - (value / count_max) * plot_height

    def y_right(value: float) -> float:
        return margin_top + plot_height - (value / effort_max) * plot_height

    bar_width = min(44, max(20, plot_width / max(label_count, 1) * 0.42))
    parts = [
        f'<div style="margin: 0.25rem 0 0.75rem 0;"><strong>{html_escape(title)}</strong></div>',
        f'<svg viewBox="0 0 {width} {height}" width="100%" height="{height}" role="img" aria-label="{html_escape(title)}">',
        '<rect x="0" y="0" width="100%" height="100%" rx="14" ry="14" fill="white" fill-opacity="0.04" stroke="rgba(255,255,255,0.12)"/>',
    ]

    # Grid and axes.
    for tick in left_ticks:
        y = y_left(tick)
        parts.append(f'<line x1="{margin_left}" y1="{y:.2f}" x2="{width - margin_right}" y2="{y:.2f}" stroke="rgba(255,255,255,0.14)" stroke-dasharray="4 4"/>')
        parts.append(f'<text x="{margin_left - 10}" y="{y + 4:.2f}" text-anchor="end" font-size="11" fill="#dbe5ff">{int(tick)}</text>')
    for tick in right_ticks:
        y = y_right(tick)
        parts.append(f'<text x="{width - margin_right + 10}" y="{y + 4:.2f}" text-anchor="start" font-size="11" fill="#dbe5ff">{int(tick)}</text>')

    parts.append(f'<line x1="{margin_left}" y1="{margin_top}" x2="{margin_left}" y2="{base_y}" stroke="rgba(255,255,255,0.55)" stroke-width="1.4"/>')
    parts.append(f'<line x1="{width - margin_right}" y1="{margin_top}" x2="{width - margin_right}" y2="{base_y}" stroke="rgba(255,255,255,0.55)" stroke-width="1.4"/>')
    parts.append(f'<line x1="{margin_left}" y1="{base_y}" x2="{width - margin_right}" y2="{base_y}" stroke="rgba(255,255,255,0.55)" stroke-width="1.4"/>')
    parts.append(f'<text x="{margin_left}" y="{18}" text-anchor="start" font-size="13" fill="#eff4ff">{html_escape(title)}</text>')
    parts.append(f'<text x="{margin_left}" y="{height - 18}" text-anchor="start" font-size="11" fill="#dbe5ff">Messstellen</text>')
    parts.append(f'<text x="{width - margin_right}" y="{height - 18}" text-anchor="end" font-size="11" fill="#dbe5ff">{html_escape(effort_label)}</text>')

    line_points = []
    for idx, (label, count, effort) in enumerate(zip(labels, counts, efforts)):
        center_x = margin_left + (idx + 0.5) * (plot_width / max(label_count, 1))
        bar_x = center_x - (bar_width / 2)
        bar_height = (count / count_max) * plot_height if count_max else 0
        bar_y = base_y - bar_height
        effort_y = y_right(effort)

        parts.append(f'<rect x="{bar_x:.2f}" y="{bar_y:.2f}" width="{bar_width:.2f}" height="{bar_height:.2f}" rx="6" ry="6" fill="#67a9ff" fill-opacity="0.88"/>')

        if mode == "Wert im Balken":
            label_text = f"{effort:.1f} h"
            text_y = bar_y + 18 if bar_height >= 24 else max(bar_y - 6, margin_top + 14)
            parts.append(f'<text x="{center_x:.2f}" y="{text_y:.2f}" text-anchor="middle" font-size="11" font-weight="600" fill="#ffffff">{html_escape(label_text)}</text>')
        else:
            line_points.append((center_x, effort_y))

        parts.append(f'<text x="{center_x:.2f}" y="{height - 34}" text-anchor="middle" font-size="10" fill="#dbe5ff">{html_escape(str(label))}</text>')
        parts.append(f'<text x="{center_x:.2f}" y="{bar_y - 6:.2f}" text-anchor="middle" font-size="11" fill="#ffffff">{int(round(count))}</text>')

    if mode == "Zweite Achse" and len(line_points) >= 2:
        path_d = "M " + " L ".join(f"{x:.2f} {y:.2f}" for x, y in line_points)
        parts.append(f'<path d="{path_d}" fill="none" stroke="#ffce5c" stroke-width="3" stroke-linejoin="round" stroke-linecap="round"/>')
        for x, y in line_points:
            parts.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="4.5" fill="#ffce5c" stroke="#1f2f52" stroke-width="1.2"/>')
    elif mode == "Zweite Achse" and len(line_points) == 1:
        x, y = line_points[0]
        parts.append(f'<circle cx="{x:.2f}" cy="{y:.2f}" r="4.5" fill="#ffce5c" stroke="#1f2f52" stroke-width="1.2"/>')

    parts.append('</svg>')
    parts.append(
        '<div style="font-size: 0.8rem; color: #dbe5ff; margin-top: 0.35rem;">'
        f'Blaue Balken = Messstellen, gelbe Linie = {html_escape(effort_label)}.'
        '</div>'
    )
    return "".join(parts)


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






def empty_override_table() -> pd.DataFrame:
    return pd.DataFrame(
        columns=[
            "Asset ID",
            "override_action",
            "override_prio",
            "override_substage",
            "override_reason",
            "override_shutdown",
        ]
    )

def normalize_override_rows(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=["Asset ID", "override_action", "override_prio", "override_substage", "override_reason", "override_shutdown"])

    out = pd.DataFrame()
    cols = {c.strip().lower(): c for c in df.columns}

    def pick(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    asset_col = pick("Asset ID", "asset_id", "assetid")
    action_col = pick("override_action", "action")
    prio_col = pick("override_prio", "prio", "priority")
    sub_col = pick("override_substage", "substage", "prio_substage")
    reason_col = pick("override_reason", "reason")
    shutdown_col = pick("override_shutdown", "unterbruch_erforderlich", "shutdown")

    if not asset_col:
        return pd.DataFrame(columns=["Asset ID", "override_action", "override_prio", "override_substage", "override_reason", "override_shutdown"])

    out["Asset ID"] = df[asset_col].astype("string").str.strip()
    out["override_action"] = df[action_col].astype("string").str.strip().str.lower() if action_col else "assign_prio"
    out["override_prio"] = df[prio_col].astype("string").str.strip().str.upper() if prio_col else "P1"
    out["override_substage"] = df[sub_col].astype("string").str.strip() if sub_col else ""
    out["override_reason"] = df[reason_col].astype("string").str.strip() if reason_col else "Manueller Override"
    if shutdown_col:
        out["override_shutdown"] = df[shutdown_col].astype("string").str.strip().str.lower().isin(["1", "true", "yes", "ja"])
    else:
        out["override_shutdown"] = False

    out = out[out["Asset ID"].notna() & (out["Asset ID"] != "")]
    return out.drop_duplicates(subset=["Asset ID"], keep="last")

def list_criteria_sets() -> List[Path]:
    return sorted(CRITERIA_DIR.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)


def default_criteria_set() -> Dict[str, Any]:
    return {
        "name": "kau_ms_relabeling_v1",
        "version": datetime.now().strftime("%Y-%m-%d_%H%M%S"),
        "notes": "KAU MS Relabeling Standardlogik",
        "config": {
            "asset_id_length_gt": 30,
            "time_scope_months_max": 5,
            "time_immediate_interval_lt": 12,
            "access_easy_patterns": [
                r"\beinfach\b",
                r"\beasy\b",
                r"\bjederzeit\b",
                r"\btechnikbereich\b",
                r"\blabor\b",
                r"\bd[-\s]?zone\b",
                r"\bleicht\b",
            ],
            "access_hard_patterns": [
                r"\bsehr\s*schwer\b",
                r"\bschwer\b",
                r"\breinraum\b",
                r"\bstillstand\b",
                r"\bzone\s*[abc]\b",
                r"\bzone[abc]\b",
                r"\babc\b",
            ],
            "qc_keywords": ["qc", "itot", "self", "selbst", "ausschluss", "excluded"],
            "qc_scope_keywords": ["qc", "itot", "self", "selbst", "ausschluss", "excluded"],
        },
        "rules": [
            {
                "id": "exclude_qc_scope",
                "active": True,
                "action": "exclude_from_prio",
                "reason": "QC/ITOT separat ausgeschlossen oder durch QC selbst zu labeln",
                "when": {"op": "==", "column": "qc_scope_status", "value": "QC_PE_excluded"},
            },
            {
                "id": "prio_p1_shutdown_scope",
                "active": True,
                "action": "assign_prio",
                "priority": "P1",
                "relabel_phase": "Shutdown_26",
                "unterbruch_erforderlich": True,
                "recommended_window": "31.07.2026-20.08.2026",
                "reason": "Lange Legacy-ID, nächste Kalibrierung im Scope, schwer zugänglich",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "IN", "column": "time_class", "values": ["T0_Immediate", "MS_Time_iScope"]},
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_krit"},
                        {"op": "==", "column": "access_class", "value": "ABC"},
                    ],
                },
            },
            {
                "id": "prio_p2_easy_scope",
                "active": True,
                "action": "assign_prio",
                "priority": "P2",
                "relabel_phase": "PostgL_bis_Oktober",
                "unterbruch_erforderlich": False,
                "recommended_window": "20.08.2026-31.10.2026",
                "reason": "Lange Legacy-ID, nächste Kalibrierung im Scope, einfach zugänglich",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "IN", "column": "time_class", "values": ["T0_Immediate", "MS_Time_iScope"]},
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_krit"},
                        {"op": "==", "column": "access_class", "value": "easy"},
                    ],
                },
            },
            {
                "id": "prio_p2_review_access",
                "active": True,
                "action": "assign_prio",
                "priority": "P2",
                "priority_substage": "review_access",
                "relabel_phase": "Klaerung_Zugaenglichkeit",
                "reason": "Lange Legacy-ID und im Scope, Zugänglichkeit unklar",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "IN", "column": "time_class", "values": ["T0_Immediate", "MS_Time_iScope"]},
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_krit"},
                        {"op": "==", "column": "access_class", "value": "unknown"},
                    ],
                },
            },
            {
                "id": "prio_p3_shutdown_27",
                "active": True,
                "action": "assign_prio",
                "priority": "P3",
                "relabel_phase": "Shutdown_27",
                "unterbruch_erforderlich": True,
                "recommended_window": "Januar 2027",
                "reason": "Nicht sofort fällig, aber schwer zugänglich und lange Legacy-ID",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "==", "column": "time_class", "value": "MS_Time_Long"},
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_krit"},
                        {"op": "==", "column": "access_class", "value": "ABC"},
                    ],
                },
            },
            {
                "id": "prio_p2a_postgl",
                "active": True,
                "action": "assign_prio",
                "priority": "P2A",
                "relabel_phase": "PostgL_bis_Januar",
                "recommended_window": "01.11.2026-31.01.2027",
                "reason": "Lange Legacy-ID, einfach zugänglich, nach Kapazität bis Januar",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "==", "column": "time_class", "value": "MS_Time_Long"},
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_krit"},
                        {"op": "==", "column": "access_class", "value": "easy"},
                    ],
                },
            },
            {
                "id": "prio_p4_opportunistic",
                "active": True,
                "action": "assign_prio",
                "priority": "P4",
                "relabel_phase": "opportunistisch_Shutdown",
                "unterbruch_erforderlich": True,
                "reason": "Legacy-ID handhabbar, aber schwer zugänglich; nur opportunistisch mitnehmen",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_iO"},
                        {"op": "==", "column": "access_class", "value": "ABC"},
                    ],
                },
            },
            {
                "id": "prio_p5_backlog",
                "active": True,
                "action": "assign_prio",
                "priority": "P5",
                "relabel_phase": "Backlog_optional",
                "reason": "Legacy-ID handhabbar und einfach zugänglich; keine Prio 1",
                "when": {
                    "op": "AND",
                    "conditions": [
                        {"op": "==", "column": "ms_legacy_class", "value": "MS_legacy_iO"},
                        {"op": "==", "column": "access_class", "value": "easy"},
                    ],
                },
            },
        ],
    }


def default_decision_tree_config() -> Dict[str, Any]:
    return json_safe(default_criteria_set()["config"])


def normalize_decision_tree_config(config: Dict[str, Any] | None) -> Dict[str, Any]:
    default_config = default_decision_tree_config()
    if not isinstance(config, dict):
        return default_config

    normalized = dict(default_config)
    for key in [
        "asset_id_length_gt",
        "time_scope_months_max",
        "time_immediate_interval_lt",
    ]:
        if key in config:
            try:
                normalized[key] = int(float(config.get(key, normalized[key])))
            except Exception:
                pass

    for key in ["access_easy_patterns", "access_hard_patterns", "qc_keywords", "qc_scope_keywords"]:
        value = config.get(key)
        if isinstance(value, str):
            items = [item.strip() for item in value.split(",") if item.strip()]
        elif isinstance(value, list):
            items = [str(item).strip() for item in value if str(item).strip()]
        else:
            items = normalized[key]
        if items:
            normalized[key] = items

    return normalized


def set_active_decision_tree_config(config: Dict[str, Any] | None) -> None:
    st.session_state["decision_tree_config"] = normalize_decision_tree_config(config)


def get_active_decision_tree_config() -> Dict[str, Any]:
    return normalize_decision_tree_config(st.session_state.get("decision_tree_config"))


def ensure_default_criteria_file() -> Path:
    default_path = CRITERIA_DIR / "kau_ms_relabeling_v1.json"
    if not default_path.exists():
        default_path.write_text(json.dumps(default_criteria_set(), ensure_ascii=False, indent=2), encoding="utf-8")
    existing = list_criteria_sets()
    if not existing:
        return default_path
    return default_path if default_path in existing else existing[0]


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

    due_col = find_column_by_candidates(working, due_candidates)
    start_col = find_column_by_candidates(working, start_candidates)
    interval_col = find_column_by_candidates(working, interval_candidates)

    today = pd.Timestamp.today().normalize()
    golive_ref = st.session_state.get("golive_reference_date", date(2026, 8, 10))
    golive_ts = pd.Timestamp(golive_ref)

    if due_col:
        due_dt = pd.to_datetime(working[due_col], errors="coerce", dayfirst=True)
        working["months_until_due"] = (due_dt - today).dt.days / 30.4375
        working["months_from_golive"] = (due_dt - golive_ts).dt.days / 30.4375
    else:
        working["months_until_due"] = pd.NA
        working["months_from_golive"] = pd.NA

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

    asset_candidates = ["Asset ID", "AssetID", "Asset_Id"]
    asset_col = find_column_by_candidates(working, asset_candidates)
    if asset_col:
        working["asset_id_length"] = working[asset_col].astype("string").str.len()
    else:
        working["asset_id_length"] = pd.NA

    return working


CSV_SCHEMA_HINTS: Dict[str, List[str]] = {
    "asset_id": ["Asset ID", "AssetID", "Asset_Id"],
    "due_date": ["Due Date BMRAM / End Datum SAP", "Due Date", "End Datum SAP"],
    "start_date": ["Start Datum SAP Auftrag", "Start Datum", "GoLive", "Go Live"],
    "interval": ["Interval", "Intervall"],
    "access_text": ["Zugänglichkeit", "Zugaenglichkeit", "Accessibility"],
    "description": ["Messstellenbeschreibung", "Messstellen Beschreibung", "Beschreibung", "Description", "Description Text"],
    "standort": ["Standort", "Location", "Site"],
    "status": ["Status"],
    "category": ["Kategorie", "Category"],
}


def _text_series_contains_any(series: pd.Series, patterns: List[str]) -> pd.Series:
    if series.empty:
        return pd.Series([False] * len(series), index=series.index)
    combined = "|".join(patterns)
    try:
        return series.str.contains(combined, case=False, regex=True, na=False)
    except re.error:
        return pd.Series([False] * len(series), index=series.index)


def classify_access_class(df: pd.DataFrame, config: Dict[str, Any] | None = None) -> pd.Series:
    config = normalize_decision_tree_config(config)
    access_col = find_column_by_candidates(df, CSV_SCHEMA_HINTS["access_text"])
    if not access_col:
        return pd.Series(["unknown"] * len(df), index=df.index, dtype="string")

    raw = df[access_col].astype("string").fillna("").str.strip()
    normalized = raw.str.lower()

    easy_mask = _text_series_contains_any(normalized, config.get("access_easy_patterns", []))
    hard_mask = _text_series_contains_any(normalized, config.get("access_hard_patterns", []))
    unknown_mask = raw.eq("") | normalized.isin(["unknown", "unbekannt", "na", "n/a", "none", "null"])

    result = pd.Series(["unknown"] * len(df), index=df.index, dtype="string")
    result.loc[hard_mask & ~easy_mask] = "ABC"
    result.loc[easy_mask & ~hard_mask] = "easy"
    result.loc[hard_mask & easy_mask] = "ABC"
    result.loc[unknown_mask] = "unknown"
    return result


def classify_legacy_class(df: pd.DataFrame, config: Dict[str, Any] | None = None) -> pd.Series:
    config = normalize_decision_tree_config(config)
    asset_col = find_column_by_candidates(df, CSV_SCHEMA_HINTS["asset_id"])
    if not asset_col:
        return pd.Series(["MS_legacy_iO"] * len(df), index=df.index, dtype="string")
    asset_len = df[asset_col].astype("string").str.len()
    result = pd.Series(["MS_legacy_iO"] * len(df), index=df.index, dtype="string")
    result.loc[asset_len > int(config.get("asset_id_length_gt", 30))] = "MS_legacy_krit"
    return result


def classify_qc_scope(df: pd.DataFrame, config: Dict[str, Any] | None = None) -> pd.Series:
    config = normalize_decision_tree_config(config)
    asset_col = find_column_by_candidates(df, CSV_SCHEMA_HINTS["asset_id"])
    qc_assets = {str(v).strip() for v in st.session_state.get("qc_green_assets", []) if str(v).strip()}
    result = pd.Series(["QC_PE_in_scope"] * len(df), index=df.index, dtype="string")
    if asset_col and qc_assets:
        asset_series = df[asset_col].astype("string").str.strip()
        result.loc[asset_series.isin(qc_assets)] = "QC_PE_excluded"

    text_cols = [c for c in [find_column_by_candidates(df, CSV_SCHEMA_HINTS["status"]), find_column_by_candidates(df, CSV_SCHEMA_HINTS["category"])] if c]
    if text_cols:
        combined = df[text_cols].astype("string").fillna("").agg(" ".join, axis=1).str.lower()
        qc_keywords = config.get("qc_scope_keywords", config.get("qc_keywords", []))
        pattern = "|".join([str(v).strip() for v in qc_keywords if str(v).strip()])
        text_mask = combined.str.contains(pattern if pattern else r"\b(qc|itot|self|selbst|ausschluss|excluded)\b", regex=True, na=False)
        result.loc[text_mask] = "QC_PE_excluded"
    return result


def find_description_column(df: pd.DataFrame) -> str | None:
    return find_column_by_candidates(df, CSV_SCHEMA_HINTS["description"])


def datalogger_mask(df: pd.DataFrame, description_col: str | None = None) -> pd.Series:
    if df is None or df.empty:
        return pd.Series([], dtype="bool")
    col = description_col if description_col and description_col in df.columns else find_description_column(df)
    if not col:
        return pd.Series([False] * len(df), index=df.index)
    series = df[col].astype("string").fillna("")
    return series.str.contains(r"datenlogger|datalogger", case=False, regex=True, na=False)


def count_datalogger_rows(df: pd.DataFrame, description_col: str | None = None) -> tuple[int, str | None]:
    mask = datalogger_mask(df, description_col=description_col)
    return int(mask.sum()), (description_col if description_col and description_col in df.columns else find_description_column(df))


def apply_datalogger_prefilter(df: pd.DataFrame, enabled: bool, description_col: str | None = None) -> tuple[pd.DataFrame, int, str | None]:
    if df is None or df.empty:
        return df, 0, description_col
    if not enabled:
        return df, 0, description_col if description_col and description_col in df.columns else find_description_column(df)

    col = description_col if description_col and description_col in df.columns else find_description_column(df)
    mask = datalogger_mask(df, description_col=col)
    return df.loc[~mask].copy(), int(mask.sum()), col


def classify_time_class(df: pd.DataFrame, golive_reference: date, config: Dict[str, Any] | None = None) -> pd.Series:
    config = normalize_decision_tree_config(config)
    due_col = find_column_by_candidates(df, CSV_SCHEMA_HINTS["due_date"])
    interval_col = find_column_by_candidates(df, CSV_SCHEMA_HINTS["interval"])
    if not due_col:
        return pd.Series(["MS_Time_Long"] * len(df), index=df.index, dtype="string")

    due_dt = pd.to_datetime(df[due_col], errors="coerce", dayfirst=True)
    interval_num = pd.to_numeric(df[interval_col], errors="coerce") if interval_col else pd.Series([pd.NA] * len(df), index=df.index)
    golive_ts = pd.Timestamp(golive_reference)
    months_from_golive = (due_dt - golive_ts).dt.days / 30.4375

    result = pd.Series(["MS_Time_Long"] * len(df), index=df.index, dtype="string")
    valid_due = due_dt.notna()

    immediate_mask = valid_due & (months_from_golive <= 0) & (pd.to_numeric(interval_num, errors="coerce") < float(config.get("time_immediate_interval_lt", 12)))
    immediate_mask = immediate_mask.fillna(False)
    scope_mask = valid_due & (months_from_golive.between(0, float(config.get("time_scope_months_max", 5)), inclusive="both")) & (pd.to_numeric(interval_num, errors="coerce") >= float(config.get("time_immediate_interval_lt", 12)))
    scope_mask = scope_mask.fillna(False)
    long_mask = valid_due & ~(immediate_mask | scope_mask)
    long_mask = long_mask.fillna(False)

    result.loc[immediate_mask] = "T0_Immediate"
    result.loc[scope_mask] = "MS_Time_iScope"
    result.loc[long_mask] = "MS_Time_Long"
    return result


def derive_decision_tree_columns(df: pd.DataFrame) -> pd.DataFrame:
    config = get_active_decision_tree_config()
    working = add_derived_time_columns(df)
    working["ms_legacy_class"] = classify_legacy_class(working, config)
    working["access_class"] = classify_access_class(working, config)
    working["qc_scope_status"] = classify_qc_scope(working, config)
    working["time_class"] = classify_time_class(working, st.session_state.get("golive_reference_date", date(2026, 8, 10)), config)
    return working


def build_csv_schema_report(df: pd.DataFrame) -> Dict[str, Any]:
    resolved = {}
    missing = []
    for key, candidates in CSV_SCHEMA_HINTS.items():
        match = find_column_by_candidates(df, candidates)
        if match:
            resolved[key] = match
        else:
            missing.append(key)

    expected = ["asset_id", "due_date", "interval", "access_text"]
    readiness_missing = [key for key in expected if key not in resolved]
    time_class_counts = df["time_class"].value_counts(dropna=False).to_dict() if "time_class" in df.columns else {}
    access_counts = df["access_class"].value_counts(dropna=False).to_dict() if "access_class" in df.columns else {}
    legacy_counts = df["ms_legacy_class"].value_counts(dropna=False).to_dict() if "ms_legacy_class" in df.columns else {}
    qc_counts = df["qc_scope_status"].value_counts(dropna=False).to_dict() if "qc_scope_status" in df.columns else {}
    description_col = resolved.get("description") or find_description_column(df)
    datalogger_count, _ = count_datalogger_rows(df, description_col=description_col)

    ready = len(readiness_missing) == 0
    return {
        "resolved": resolved,
        "missing": missing,
        "readiness_missing": readiness_missing,
        "ready": ready,
        "description_col": description_col,
        "datalogger_count": datalogger_count,
        "time_class_counts": time_class_counts,
        "access_counts": access_counts,
        "legacy_counts": legacy_counts,
        "qc_counts": qc_counts,
    }


def render_csv_check(df: pd.DataFrame) -> None:
    main_col, help_col = st.columns([4, 1.35])
    report = build_csv_schema_report(df)
    with main_col:
        st.markdown("### CSV-Check")
        ready_label = "bereit" if report["ready"] else "teilweise unvollständig"
        st.caption(f"Auto-Pipeline Status: {ready_label}")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Zeilen", f"{len(df):,}".replace(",", "'"))
        c2.metric("Erkannte Pflichtfelder", f"{len(report['resolved']) - len(report['missing'])}")
        c3.metric("Fehlende Pflichtfelder", f"{len(report['readiness_missing'])}")
        c4.metric("QC Excluded", f"{int((df.get('qc_scope_status', pd.Series(dtype='string')) == 'QC_PE_excluded').sum())}" if "qc_scope_status" in df.columns else "0")
        d1, d2 = st.columns(2)
        d1.metric("Datenlogger", f"{report['datalogger_count']}")
        d2.caption(f"Spalte: {report['description_col'] or 'nicht erkannt'}")
        prefilter_on = bool(st.session_state.get("datalogger_prefilter_on", False))
        prefilter_removed = int(st.session_state.get("datalogger_prefilter_removed_count", 0))
        st.caption(f"Datenlogger-Vorfilter: {'aktiv' if prefilter_on else 'inaktiv'} | entfernt: {prefilter_removed}")

        st.markdown("#### Erkannte Spalten")
        mapping_rows = []
        for key, label in [
            ("asset_id", "Asset-ID"),
            ("due_date", "Due Date"),
            ("start_date", "Start Datum"),
            ("interval", "Interval"),
            ("access_text", "Zugänglichkeit"),
            ("description", "Messstellenbeschreibung"),
            ("standort", "Standort"),
            ("status", "Status"),
            ("category", "Kategorie"),
        ]:
            mapping_rows.append(
                {
                    "kanonisch": label,
                    "erkannt": report["resolved"].get(key, ""),
                    "status": "OK" if key in report["resolved"] else "fehlt",
                }
            )
        mapping_df = pd.DataFrame(mapping_rows)
        st.dataframe(mapping_df, use_container_width=True, height=260, column_config=build_column_config(mapping_df, allow_manual_edit=False))

        if report["readiness_missing"]:
            st.warning("Pflichtfelder fehlen für eine saubere Auto-Planung: " + ", ".join(report["readiness_missing"]))
        else:
            st.success("Die CSV ist für die automatische Entscheidungskette grundsätzlich verwendbar.")

        st.markdown("#### Klassifizierungen")
        class_df = pd.DataFrame(
            {
                "time_class": pd.Series(report["time_class_counts"]),
                "access_class": pd.Series(report["access_counts"]),
                "ms_legacy_class": pd.Series(report["legacy_counts"]),
                "qc_scope_status": pd.Series(report["qc_counts"]),
            }
        ).fillna(0)
        st.dataframe(class_df, use_container_width=True, height=260, column_config=build_column_config(class_df, allow_manual_edit=False))
    with help_col:
        with st.expander("Worauf geprüft wird", expanded=True):
            st.caption("- Pflichtspalten werden gegen einen kanonischen Spaltenvertrag geprüft.")
            st.caption("- Daraus werden die Entscheidungsbaum-Spalten abgeleitet.")
            st.caption("- Erst danach ist der Phasenplan belastbar.")
            st.caption("- Wenn Felder fehlen, bleibt die Logik sichtbar, aber der Plan ist nur teilweise automatisiert.")


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

    if op == "IN":
        values = cond.get("values", [])
        if not isinstance(values, list):
            values = [values]
        left_text = df[col].astype("string").str.strip()
        norm_vals = [str(v).strip() for v in values]
        return left_text.isin(norm_vals)

    if op == "IN_REF_LIST":
        ref_name = str(cond.get("ref_list", "")).strip()
        ref_values = reference_lists.get(ref_name, set())
        if not ref_values:
            return pd.Series([False] * len(df), index=df.index)
        return df[col].astype("string").str.strip().isin(ref_values)

    return pd.Series([False] * len(df), index=df.index)




def coerce_when_to_dict(raw_when: Any) -> Dict[str, Any] | None:
    if isinstance(raw_when, dict):
        return raw_when
    if isinstance(raw_when, str):
        s = raw_when.strip()
        if not s:
            return None
        # Try JSON first, then Python literal format from data_editor export.
        try:
            loaded = json.loads(s)
            if isinstance(loaded, dict):
                return loaded
        except Exception:
            pass
        try:
            loaded = ast.literal_eval(s)
            if isinstance(loaded, dict):
                return loaded
        except Exception:
            pass
    return None


def normalize_rule(rule: Dict[str, Any]) -> Dict[str, Any]:
    normalized = dict(rule)
    when_obj = coerce_when_to_dict(normalized.get("when"))
    if when_obj is not None:
        normalized["when"] = when_obj
    if isinstance(normalized.get("mark_shutdown"), str):
        normalized["mark_shutdown"] = normalized.get("mark_shutdown", "").strip().lower() in {"1", "true", "yes", "ja"}
    return normalized

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
    prio_stage = pd.Series(["P6"] * len(working), index=working.index, dtype="string")
    prio_substage = pd.Series(["" ] * len(working), index=working.index, dtype="string")
    prio_rank = pd.Series([6.0] * len(working), index=working.index, dtype="float64")
    shutdown_required = pd.Series([False] * len(working), index=working.index)
    relabel_phase = pd.Series(["" for _ in range(len(working))], index=working.index, dtype="string")
    recommended_window = pd.Series(["" for _ in range(len(working))], index=working.index, dtype="string")
    decision_reason = pd.Series(["" for _ in range(len(working))], index=working.index, dtype="string")

    for raw_rule in criteria.get("rules", []):
        rule = normalize_rule(raw_rule)
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
            target_rank = {"P1": 1.0, "P2": 2.0, "P2A": 2.5, "P3": 3.0, "P4": 4.0, "P5": 5.0, "P6": 6.0}.get(target, 6.0)
            upgrade = mask.fillna(False) & (target_rank < prio_rank)
            prio_rank.loc[upgrade] = target_rank
            prio_stage.loc[upgrade] = target
            rule_col.loc[upgrade] = (rule_col.loc[upgrade] + "; " + rid).str.strip("; ")
            if bool(rule.get("mark_shutdown", False)) or bool(rule.get("unterbruch_erforderlich", False)):
                shutdown_required.loc[mask.fillna(False)] = True
            substage = str(rule.get("priority_substage", rule.get("substage", ""))).strip()
            if substage:
                prio_substage.loc[mask.fillna(False)] = substage
            phase = str(rule.get("relabel_phase", rule.get("phase", ""))).strip()
            if phase:
                relabel_phase.loc[upgrade] = phase
            window = str(rule.get("recommended_window", "")).strip()
            if window:
                recommended_window.loc[upgrade] = window
            if reason:
                decision_reason.loc[upgrade] = reason

    excluded = working[exclude_mask].copy()
    excluded["decision_status"] = "excluded"
    excluded["exclude_reasons"] = reason_col.loc[exclude_mask].fillna("")
    excluded["matched_rule_ids"] = rule_col.loc[exclude_mask].fillna("")
    excluded["prio_stage"] = "excluded"
    excluded["prio_substage"] = ""
    excluded["relabel_phase"] = "excluded"
    excluded["recommended_window"] = ""
    excluded["unterbruch_erforderlich"] = shutdown_required.loc[exclude_mask].fillna(False)
    excluded["decision_reason"] = reason_col.loc[exclude_mask].fillna("")

    included = working[~exclude_mask].copy()
    included["decision_status"] = "in_prio"
    included["exclude_reasons"] = ""
    included["matched_rule_ids"] = rule_col.loc[~exclude_mask].fillna("")
    included["prio_stage"] = prio_stage.loc[~exclude_mask].fillna("P6")
    included["prio_substage"] = prio_substage.loc[~exclude_mask].fillna("")
    included["relabel_phase"] = relabel_phase.loc[~exclude_mask].fillna("")
    included["recommended_window"] = recommended_window.loc[~exclude_mask].fillna("")
    included["unterbruch_erforderlich"] = shutdown_required.loc[~exclude_mask].fillna(False)
    included["decision_reason"] = decision_reason.loc[~exclude_mask].fillna("")
    return included, excluded


def _find_rule(criteria: Dict[str, Any], rule_id: str) -> Dict[str, Any] | None:
    for raw_rule in criteria.get("rules", []):
        rule = normalize_rule(raw_rule)
        if str(rule.get("id", "")) == rule_id:
            return rule
    return None


def _count_rule_matches(df: pd.DataFrame, rule: Dict[str, Any], reference_lists: Dict[str, set]) -> int:
    rule = normalize_rule(rule)
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
    with st.sidebar.expander("Live-Entscheidungsbaum", expanded=True):
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

    with st.sidebar.expander("Datenlogger-Info", expanded=False) as dl_box:
        description_candidates = [c for c in text_cols if any(k in c.lower() for k in ["beschreibung", "description", "desc", "messstellen"])]
        if not description_candidates:
            description_candidates = text_cols
        if description_candidates:
            default_description = find_description_column(df) if find_description_column(df) in description_candidates else description_candidates[0]
            description_idx = description_candidates.index(default_description) if default_description in description_candidates else 0
            description_col = dl_box.selectbox(
                "Messstellenbeschreibung-Spalte",
                description_candidates,
                index=description_idx,
                key=f"{prefix}_datalogger_desc_col",
            )
            dl_mask = datalogger_mask(filtered, description_col=description_col)
            dl_count = int(dl_mask.sum())
            dl_box.caption(f"Datenlogger gefunden in aktueller Ansicht: {dl_count}")
            dl_box.caption("Der eigentliche Ausschluss läuft als globaler Vorfilter in der Sidebar.")
        else:
            dl_box.caption("Keine Textspalte für die Messstellenbeschreibung verfügbar.")

    filtered, excluded_qc = apply_qc_exclusion(filtered, prefix, st.session_state.get)
    st.session_state[f"{prefix}_qc_excluded_df"] = excluded_qc

    criteria = st.session_state.get("active_criteria_set", {})
    reference_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}
    filtered, excluded_rules = apply_criteria_rules(filtered, criteria, reference_lists)
    st.session_state[f"{prefix}_criteria_excluded_df"] = excluded_rules

    sync_store_from_session(prefix)
    touch_view_timestamp(prefix)
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



def _short_list(vals: List[Any], max_items: int = 3) -> str:
    items = [str(v) for v in vals if str(v).strip()]
    if not items:
        return ""
    if len(items) <= max_items:
        return ", ".join(items)
    return ", ".join(items[:max_items]) + f" (+{len(items)-max_items})"


def build_view_summary(df: pd.DataFrame, prefix: str) -> str:
    state = get_filter_store(prefix)
    parts: List[str] = []

    inc = str(state.get(f"{prefix}_regex_include", "")).strip()
    exc = str(state.get(f"{prefix}_regex_exclude", "")).strip()
    rcol = str(state.get(f"{prefix}_regex_col", "")).strip()
    if rcol and (inc or exc):
        if inc:
            parts.append(f"Regex+ {rcol}: {inc}")
        if exc:
            parts.append(f"Regex- {rcol}: {exc}")

    text_cols = get_text_columns(df)
    cat_hits = []
    for col in text_cols:
        sel = state.get(f"{prefix}_cat_{col}", [])
        if isinstance(sel, list) and sel:
            cat_hits.append(f"{col}={_short_list(sel, 2)}")
    if cat_hits:
        parts.append("Kat " + " | ".join(cat_hits[:3]))

    ncols = state.get(f"{prefix}_num_cols", [])
    if isinstance(ncols, list) and ncols:
        num_desc = []
        for col in ncols[:3]:
            rng = state.get(f"{prefix}_num_rng_{col}")
            if isinstance(rng, (list, tuple)) and len(rng) == 2:
                num_desc.append(f"{col}:[{rng[0]}..{rng[1]}]")
            else:
                num_desc.append(str(col))
        parts.append("Num " + " | ".join(num_desc))

    dcol = state.get(f"{prefix}_date_col")
    dr = state.get(f"{prefix}_date_range")
    if dcol and dcol != "(kein Filter)" and isinstance(dr, (list, tuple)) and len(dr) == 2:
        parts.append(f"Datum {dcol}: {dr[0]}..{dr[1]}")

    criteria = st.session_state.get("active_criteria_set", {})
    cname = str(criteria.get("name", ""))
    if cname and cname != "none":
        parts.append(f"Entscheidungsbaum: {cname}")
    else:
        parts.append("Entscheidungsbaum: keine")

    return " | ".join(parts) if parts else "Keine aktiven Filter"


def touch_view_timestamp(prefix: str) -> None:
    st.session_state[f"{prefix}_last_changed"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def render_view(df: pd.DataFrame, view_name: str, prefix: str, active_prefix: str) -> None:
    st.markdown(f"### {view_name}")
    cmeta1, cmeta2 = st.columns([3, 2])
    with cmeta1:
        purpose = st.text_input(
            "Ansichtszweck",
            value=st.session_state.get(f"{prefix}_purpose", ""),
            placeholder="z.B. P1A-Unterbruch Fokus Verpackung_PE",
            key=f"{prefix}_purpose",
        )
    with cmeta2:
        st.text_input(
            "Filter-Zusammenfassung (auto)",
            value=build_view_summary(df, prefix),
            key=f"{prefix}_summary_preview",
            disabled=True,
        )
    if f"{prefix}_last_changed" not in st.session_state:
        touch_view_timestamp(prefix)
    st.caption(f"Zuletzt geändert: {st.session_state.get(f'{prefix}_last_changed','-')}")
    if st.button("Diese Ansicht links bearbeiten", key=f"{prefix}_activate_sidebar"):
        st.session_state["pending_active_prefix"] = prefix
        st.rerun()
    if active_prefix not in {"", "__none__", prefix}:
        st.caption("Hinweis: Sidebar-Filter sind aktuell auf eine andere Ansicht gesetzt.")
    if st.session_state.pop(f"{prefix}_preset_loaded_msg", False):
        st.success("Preset geladen. Ansicht aktualisiert.")

    filtered = apply_filters(df, prefix, show_sidebar_filters=(prefix == active_prefix))
    render_view_preset_tools(prefix, view_name)
    st.write(f"{len(filtered)} von {len(df)} Zeilen")
    edited = apply_manual_override(filtered, prefix)
    render_summary(edited, prefix)

    if "prio_stage" in edited.columns:
        st.markdown("#### Prioritätsstufen (aus Entscheidungsbaum)")
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
        st.markdown("#### Ausgeschlossen durch Entscheidungsbaum")
        st.write(f"{len(excluded_criteria)} Zeilen wurden durch den aktiven Entscheidungsbaum ausgeschlossen.")
        default_cols = [c for c in ["Asset ID", "Gebäude / MU", "Standort", "matched_rule_ids", "exclude_reasons"] if c in excluded_criteria.columns]
        sel_cols = st.multiselect(
            "Spalten in Ausschlussliste",
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






def upsert_rule(criteria: Dict[str, Any], rule: Dict[str, Any]) -> None:
    rid = str(rule.get("id", "")).strip()
    rules = criteria.setdefault("rules", [])
    for i, existing in enumerate(rules):
        if str(existing.get("id", "")).strip() == rid:
            rules[i] = rule
            return
    rules.append(rule)


def delete_rule_by_id(criteria: Dict[str, Any], rule_id: str) -> None:
    rules = criteria.get("rules", [])
    criteria["rules"] = [r for r in rules if str(r.get("id", "")).strip() != str(rule_id).strip()]

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
    st.markdown("#### Entscheidungsbaum-Builder")
    st.caption("Neu: beliebig viele Bedingungen je Regel (AND/OR), inkl. REGEX und optionalem NOT.")

    columns = list(df.columns) + ["months_until_due", "months_since_start", "interval_half_months", "months_from_golive", "asset_id_length"]
    if not columns:
        st.info("Keine Spalten verfügbar.")
        return

    c1, c2, c3 = st.columns([1, 1, 1])
    action = c1.selectbox("Aktion", ["exclude_from_prio", "assign_prio"], key="rb_action")
    priority = c2.selectbox("Prio (bei assign)", ["P1", "P2", "P3"], key="rb_priority")
    logic = c3.selectbox("Verknüpfung", ["AND", "OR"], key="rb_logic")

    count = st.number_input("Anzahl Bedingungen", min_value=1, max_value=8, value=3, step=1, key="rb_cond_count")
    conditions = []
    operators = [">", ">=", "<", "<=", "==", "!=", "REGEX"]

    for i in range(int(count)):
        with st.expander(f"Bedingung {i+1}", expanded=(i < 2)):
            a0, a1, a2, a3, a4 = st.columns([1, 2, 1, 1, 2])
            negate = a0.checkbox("NOT", value=False, key=f"rb_not_{i}")
            col = a1.selectbox("Spalte", columns, key=f"rb_col_{i}")
            op = a2.selectbox("Operator", operators, key=f"rb_op_{i}")

            if op == "REGEX":
                pattern = a4.text_input("Regex", value="", key=f"rb_regex_{i}")
                cond_core = {"op": "REGEX", "column": col, "pattern": pattern}
            else:
                mode = a3.selectbox("Werttyp", ["Konstante", "Spalte"], key=f"rb_mode_{i}")
                if mode == "Konstante":
                    raw_val = a4.text_input("Wert", value="0", key=f"rb_val_{i}")
                    try:
                        val = float(raw_val)
                    except Exception:
                        val = raw_val
                    cond_core = {"op": op, "column": col, "value": val}
                else:
                    val_col = a4.selectbox("Vergleichsspalte", columns, key=f"rb_val_col_{i}")
                    cond_core = {"op": op, "column": col, "value_col": val_col}

            cond = {"op": "NOT", "condition": cond_core} if negate else cond_core
            conditions.append(cond)

    rule_id = st.text_input("Regel-ID", value=f"rule_{datetime.now().strftime('%H%M%S')}", key="rb_id")
    reason = st.text_input("Begründung", value="Manuell erstellte Regel", key="rb_reason")

    when = {"op": logic, "conditions": conditions}
    st.code(json.dumps(when, ensure_ascii=False, indent=2), language="json")
    st.caption(f"Vorschau: {build_condition_preview(when)}")

    if st.button("Regel zum aktiven Entscheidungsbaum hinzufügen", key="rb_add"):
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
        upsert_rule(criteria, new_rule)
        st.session_state["active_criteria_set"] = criteria
        st.success("Regel hinzugefügt/aktualisiert. Jetzt unten 'Entscheidungsbaum speichern' klicken.")
        st.rerun()



def render_prio34_shutdown_recipe(df: pd.DataFrame) -> None:
    st.markdown("#### Rezept: P1 (Länge) + P1A (Unterbruch)")
    asset_col = find_column_by_candidates(df, ["Asset ID", "AssetID", "Asset_Id"])
    access_col = find_column_by_candidates(df, ["Zugänglichkeit", "Zugaenglichkeit", "Accessibility"])

    c1, c2, c3 = st.columns([1, 1, 2])
    len_threshold = c1.number_input("Zeichen-Schwelle", min_value=1, max_value=200, value=34, step=1, key="recipe_len_threshold")
    include_mittel = c2.checkbox("'mittel' mit einschließen", value=False, key="recipe_include_mittel")
    substage = c3.text_input("Sub-Prio für Unterbruch", value="1A", key="recipe_substage")

    default_rx = "schwer|sehr\s*schwer"
    if include_mittel:
        default_rx = "mittel|schwer|sehr\s*schwer"
    access_regex = st.text_input("Regex für unterbruchspflichtige Zugänglichkeit", value=default_rx, key="recipe_access_regex")

    st.caption("Regel A: asset_id_length > Schwelle => P1")
    st.caption("Regel B: asset_id_length <= Schwelle UND Zugänglichkeit trifft Regex => P1 + Subprio + Unterbruch")

    if st.button("Rezept-Regeln hinzufügen/aktualisieren", key="recipe_apply_btn"):
        if not asset_col:
            st.error("Asset-ID-Spalte nicht gefunden.")
            return
        if not access_col:
            st.error("Zugänglichkeits-Spalte nicht gefunden.")
            return
        try:
            re.compile(access_regex)
        except re.error as exc:
            st.error(f"Ungültiger Regex: {exc}")
            return

        criteria = st.session_state.get("active_criteria_set", default_criteria_set())

        rule_a = {
            "id": "prio1_asset_len_gt_threshold",
            "active": True,
            "action": "assign_prio",
            "priority": "P1",
            "reason": f"Asset-ID länger als {int(len_threshold)} Zeichen",
            "when": {"op": ">", "column": "asset_id_length", "value": float(len_threshold)},
        }

        rule_b = {
            "id": "prio1a_short_but_hard_access",
            "active": True,
            "action": "assign_prio",
            "priority": "P1",
            "priority_substage": str(substage).strip() or "1A",
            "mark_shutdown": True,
            "reason": f"<= {int(len_threshold)} Zeichen, aber schwer zugänglich: Relabeling nur bei Unterbruch",
            "when": {
                "op": "AND",
                "conditions": [
                    {"op": "<=", "column": "asset_id_length", "value": float(len_threshold)},
                    {"op": "REGEX", "column": access_col, "pattern": access_regex},
                ],
            },
        }

        upsert_rule(criteria, rule_a)
        upsert_rule(criteria, rule_b)
        st.session_state["active_criteria_set"] = criteria
        st.success("Rezept-Regeln übernommen. Danach Entscheidungsbaum speichern.")
        st.rerun()

    if st.button("Rezept-Regeln entfernen", key="recipe_remove_btn"):
        criteria = st.session_state.get("active_criteria_set", default_criteria_set())
        delete_rule_by_id(criteria, "prio1_asset_len_gt_threshold")
        delete_rule_by_id(criteria, "prio1a_short_but_hard_access")
        delete_rule_by_id(criteria, "shutdown_override_access")
        delete_rule_by_id(criteria, "shutdown_override_access_length")
        st.session_state["active_criteria_set"] = criteria
        st.success("Rezept-/Unterbruch-Regeln entfernt. Danach Entscheidungsbaum speichern.")
        st.rerun()

def render_golive_hint() -> None:
    st.markdown("#### Go-Live Regelhilfe")
    ref = st.session_state.get("golive_reference_date", date(2026, 8, 10))
    st.caption(f"Aktuelles Go-Live Referenzdatum: {ref.isoformat()}")
    st.caption("Für deinen Fall im Builder setzen: Interval < 12 UND months_from_golive >= 0 UND months_from_golive <= 5")


TAB_HELP: Dict[str, Dict[str, List[str]]] = {
    "criteria": {
        "Was ist der Zweck?": [
            "Hier wird die offizielle Entscheidungslogik gepflegt.",
            "Regeln bleiben erst dauerhaft erhalten, wenn der Entscheidungsbaum gespeichert wird.",
        ],
        "Wichtige Begriffe": [
            "assign_prio: Zeile bleibt drin und bekommt eine Prio.",
            "exclude_from_prio: Zeile wird aus der Prio-Liste entfernt und in Ausschlüsse gezeigt.",
            "when: Bedingung der Regel, z. B. AND/OR/NOT mit Spaltenvergleichen.",
        ],
        "Typischer Ablauf": [
            "Eine Regelidee ändern, Regel hinzufügen, neue Version speichern, dann im Vergleich prüfen.",
            "Für Ausnahmen einzelne Assets besser über Manuelle Overrides pflegen.",
        ],
    },
    "decision_tree": {
        "Was sehe ich hier?": [
            "Die Reihenfolge der automatischen Klassifikation von Scope bis Prio.",
            "Diese Stufen erklären, warum eine Messstelle P1, P2A, P4 usw. wird.",
        ],
        "Abkürzungen": [
            "ms_legacy_class: Einordnung über Asset-ID-Länge.",
            "MS_legacy_krit: Asset ID länger als 30 Zeichen.",
            "MS_legacy_iO: Asset ID bis 30 Zeichen.",
            "access_class ABC: schwer/Zone/Reinraum/Stillstand oder ähnliche Hinweise.",
            "access_class easy: einfach/jederzeit/Technikbereich/Labor/D-Zone.",
            "time_class: zeitliche Lage aus Due Date, Intervall und Go-Live.",
        ],
    },
    "phase_plan": {
        "Was sehe ich hier?": [
            "Mengen und Aufwand je Relabeling-Phase und Prio.",
            "Die Berechnung nutzt die Kapazitätswerte aus der Sidebar.",
        ],
        "So lesen": [
            "Personentage zeigen Aufwand mit Puffer und Teamgröße.",
            "Wochenbedarf rechnet Personentage in Arbeitswochen um.",
        ],
    },
    "prio_matrix": {
        "Was sehe ich hier?": [
            "Eine Kreuztabelle aus Zeitklasse, Zugänglichkeit und Legacy-Klasse.",
            "Gut zur Plausibilitätsprüfung vor Diskussionen mit Fachbereichen.",
        ],
        "Prüffragen": [
            "Sind viele unknown-Zugänglichkeiten vorhanden?",
            "Gibt es viele kritische Legacy-IDs außerhalb des Zeitfensters?",
        ],
    },
    "exclusions": {
        "Was sehe ich hier?": [
            "Alle Messstellen, die durch den Entscheidungsbaum ausgeschlossen wurden.",
            "Wichtig sind exclude_reasons und matched_rule_ids.",
        ],
        "Prüffragen": [
            "Ist der Ausschlussgrund fachlich verständlich?",
            "Sind QC/ITOT-Ausschlüsse korrekt separiert?",
        ],
    },
}


def render_context_help(help_key: str) -> None:
    content = TAB_HELP.get(help_key, {})
    st.markdown("#### Hilfe zu diesem Tab")
    if not content:
        st.caption("Keine tab-spezifische Hilfe hinterlegt.")
        return
    for title, lines in content.items():
        with st.expander(title, expanded=True):
            for line in lines:
                st.caption(f"- {line}")


def render_criteria_editor(df: pd.DataFrame) -> None:
    main_col, help_col = st.columns([4, 1.35])
    with main_col:
        st.markdown("### Entscheidungsbaum")
        render_golive_hint()
        render_prio34_shutdown_recipe(df)
        st.markdown("#### Entscheidungsbaum-Konfiguration")
        cfg = get_active_decision_tree_config()
        c1, c2, c3 = st.columns(3)
        cfg_asset = c1.number_input(
            "Legacy-Schwelle Asset-ID >",
            min_value=1,
            max_value=120,
            value=int(cfg.get("asset_id_length_gt", 30)),
            step=1,
            key="dt_asset_id_length_gt",
        )
        cfg_scope = c2.number_input(
            "iScope Fenster (Monate)",
            min_value=0,
            max_value=24,
            value=int(cfg.get("time_scope_months_max", 5)),
            step=1,
            key="dt_time_scope_months_max",
        )
        cfg_interval = c3.number_input(
            "Immediate wenn Intervall <",
            min_value=1,
            max_value=36,
            value=int(cfg.get("time_immediate_interval_lt", 12)),
            step=1,
            key="dt_time_immediate_interval_lt",
        )
        c4, c5 = st.columns(2)
        easy_patterns = c4.text_area(
            "Zugänglichkeit easy Muster (kommagetrennt)",
            value=", ".join(cfg.get("access_easy_patterns", [])),
            key="dt_access_easy_patterns",
            height=90,
        )
        hard_patterns = c5.text_area(
            "Zugänglichkeit hard Muster (kommagetrennt)",
            value=", ".join(cfg.get("access_hard_patterns", [])),
            key="dt_access_hard_patterns",
            height=90,
        )
        qc_patterns = st.text_area(
            "QC/ITOT Muster (kommagetrennt)",
            value=", ".join(cfg.get("qc_keywords", [])),
            key="dt_qc_keywords",
            height=80,
        )
        current_cfg = {
            "asset_id_length_gt": int(cfg_asset),
            "time_scope_months_max": int(cfg_scope),
            "time_immediate_interval_lt": int(cfg_interval),
            "access_easy_patterns": [p.strip() for p in easy_patterns.split(",") if p.strip()],
            "access_hard_patterns": [p.strip() for p in hard_patterns.split(",") if p.strip()],
            "qc_keywords": [p.strip() for p in qc_patterns.split(",") if p.strip()],
            "qc_scope_keywords": [p.strip() for p in qc_patterns.split(",") if p.strip()],
        }
        set_active_decision_tree_config(current_cfg)
        render_rule_builder(df)
        criteria = st.session_state.get("active_criteria_set", default_criteria_set())
        st.write(f"Aktiv: {criteria.get('name', '')} | Version: {criteria.get('version', '')}")
        if criteria.get("name") == "none":
            st.info("Kein Entscheidungsbaum aktiv. Es werden nur manuelle Filter angewendet.")
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
        if c2.button("Entscheidungsbaum speichern", key="criteria_save_btn"):
            filename_raw = st.session_state.get("criteria_save_name", "").strip()
            stem_name = Path(filename_raw).stem if filename_raw else "criteria_set"
            current_name = str(criteria.get("name", "")).strip()
            effective_name = stem_name if current_name in {"", "none"} else current_name

            new_set = {
                "name": effective_name,
                "version": datetime.now().strftime("%Y-%m-%d_%H%M%S"),
                "notes": st.session_state.get("criteria_notes", ""),
                "config": get_active_decision_tree_config(),
                "rules": edited_rules.fillna("").to_dict(orient="records"),
            }
            target = save_criteria(new_set, filename)
            st.session_state["active_criteria_set"] = new_set
            st.session_state["active_criteria_path"] = str(target)
            st.success(f"Gespeichert und aktiviert: {target}")

        st.caption("Regeltypen: numeric_gt, string_length_lt, regex_match, date_between, ref_list_match oder when-Block mit AND/OR/NOT")
        st.caption("Action: exclude_from_prio oder assign_prio (priority: P1/P2/P3)")

        with st.expander("Aktiver Entscheidungsbaum (JSON Vorschau)", expanded=False):
            st.code(json.dumps(criteria, ensure_ascii=False, indent=2), language="json")

        st.markdown("#### Regel entfernen")
        ids = [str(r.get("id", "")) for r in criteria.get("rules", []) if str(r.get("id", "")).strip()]
        if ids:
            col_del_1, col_del_2 = st.columns([3, 1])
            rid = col_del_1.selectbox("Regel-ID", ids, key="criteria_delete_rule_id")
            if col_del_2.button("Regel löschen", key="criteria_delete_rule_btn"):
                delete_rule_by_id(criteria, rid)
                st.session_state["active_criteria_set"] = criteria
                st.success(f"Regel gelöscht: {rid}. Danach Entscheidungsbaum speichern.")
                st.rerun()
    with help_col:
        render_context_help("criteria")






def render_decision_tree(df: pd.DataFrame) -> None:
    main_col, help_col = st.columns([4, 1.35])
    with main_col:
        st.markdown("### Entscheidungsbaum")
        cfg = get_active_decision_tree_config()
        c_cfg1, c_cfg2, c_cfg3 = st.columns(3)
        c_cfg1.metric("Legacy-Schwelle", int(cfg.get("asset_id_length_gt", 30)))
        c_cfg2.metric("iScope Monate", int(cfg.get("time_scope_months_max", 5)))
        c_cfg3.metric("Immediate Intervall <", int(cfg.get("time_immediate_interval_lt", 12)))
        st.markdown(
            """
1. Scope bestimmen: `qc_scope_status`
2. QC-Ausschluss: `QC_PE_excluded` -> Ausschluss
3. Zeitklasse: `time_class` (`T0_Immediate`, `MS_Time_iScope`, `MS_Time_Long`)
4. Legacy-Klasse: `ms_legacy_class` (`MS_legacy_krit`/`MS_legacy_iO`)
5. Zugänglichkeit: `access_class` (`ABC`/`easy`/`unknown`)
6. Ergebnis: `prio_stage`, `prio_substage`, `relabel_phase`, `recommended_window`, `unterbruch_erforderlich`
            """
        )
        if {"time_class", "access_class", "ms_legacy_class", "qc_scope_status"}.issubset(df.columns):
            st.markdown("#### Aktuelle Klassifikation")
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("T0_Immediate", int((df["time_class"] == "T0_Immediate").sum()))
            c2.metric("MS_Time_iScope", int((df["time_class"] == "MS_Time_iScope").sum()))
            c3.metric("MS_Time_Long", int((df["time_class"] == "MS_Time_Long").sum()))
            c4.metric("QC Excluded", int((df["qc_scope_status"] == "QC_PE_excluded").sum()))
            class_overview = pd.DataFrame(
                {
                    "time_class": df["time_class"].value_counts(dropna=False),
                    "access_class": df["access_class"].value_counts(dropna=False),
                    "ms_legacy_class": df["ms_legacy_class"].value_counts(dropna=False),
                }
            ).fillna(0)
            st.dataframe(class_overview, use_container_width=True, height=240, column_config=build_column_config(class_overview, allow_manual_edit=False))
    with help_col:
        render_context_help("decision_tree")


def render_prio_matrix(df: pd.DataFrame) -> None:
    main_col, help_col = st.columns([4, 1.35])
    with main_col:
        st.markdown("### Prio-Matrix")
        cols = ["time_class", "access_class", "ms_legacy_class"]
        if not all(c in df.columns for c in cols):
            st.info("Benötigte Klassenspalten fehlen.")
        else:
            matrix = pd.pivot_table(df, index=["time_class", "access_class"], columns="ms_legacy_class", aggfunc="size", fill_value=0)
            st.dataframe(matrix, use_container_width=True, height=420)
    with help_col:
        render_context_help("prio_matrix")


def render_exclusions(df_ex: pd.DataFrame) -> None:
    main_col, help_col = st.columns([4, 1.35])
    with main_col:
        st.markdown("### Ausschlüsse")
        if df_ex is None or df_ex.empty:
            st.info("Keine Ausschlüsse im aktuellen Entscheidungsbaum.")
        else:
            cols = [c for c in ["Asset ID", "Gebäude / MU", "Standort", "exclude_reasons", "matched_rule_ids"] if c in df_ex.columns]
            view = df_ex[cols] if cols else df_ex
            st.dataframe(view, use_container_width=True, height=500, column_config=build_column_config(view, allow_manual_edit=False))
    with help_col:
        render_context_help("exclusions")


def render_phase_plan(df_in: pd.DataFrame, source_path: str | None = None) -> None:
    main_col, help_col = st.columns([4, 1.35])
    with main_col:
        st.markdown("### Phasenplan")
        if df_in is None or df_in.empty or "relabel_phase" not in df_in.columns:
            st.info("Keine priorisierten Daten für Phasenplan vorhanden.")
        else:
            min_normal = float(st.session_state.get("cap_min_normal", 7))
            min_medium = float(st.session_state.get("cap_min_medium", 12))
            buffer_pct = float(st.session_state.get("cap_buffer_pct", 30))
            persons = float(st.session_state.get("cap_persons", 2))
            hours_day = float(st.session_state.get("cap_hours_day", 7))

            grp = df_in.groupby(["relabel_phase", "prio_stage"], dropna=False).size().reset_index(name="anzahl_messstellen")
            grp["aufwand_7min"] = grp["anzahl_messstellen"] * min_normal
            grp["aufwand_12min"] = grp["anzahl_messstellen"] * min_medium
            grp["aufwand_plus_puffer_min"] = grp["aufwand_12min"] * (1 + buffer_pct / 100.0)
            grp["personentage"] = grp["aufwand_plus_puffer_min"] / (persons * hours_day * 60.0)
            grp["wochenbedarf"] = grp["personentage"] / 5.0
            for col in ["aufwand_7min", "aufwand_12min", "aufwand_plus_puffer_min"]:
                grp[col] = grp[col].round(0).astype(int)
            grp["personentage"] = grp["personentage"].round(1)
            grp["wochenbedarf"] = grp["wochenbedarf"].round(1)
            csum1, csum2, csum3 = st.columns(3)
            csum1.metric("Phasen", grp["relabel_phase"].nunique())
            csum2.metric("Messstellen", int(grp["anzahl_messstellen"].sum()))
            csum3.metric("Gesamt-Personentage", f"{grp['personentage'].sum():.1f}")
            st.dataframe(grp, use_container_width=True, height=420, column_config=build_column_config(grp, allow_manual_edit=False))

            st.markdown("#### Grafischer Überblick")
            overview = build_phase_overview_tables(grp)
            if overview:
                view_mode = st.radio(
                    "Grafikansicht",
                    ["Nach Prio", "Nach Phase", "Phase x Prio"],
                    horizontal=True,
                    key="phase_overview_mode",
                )
                if view_mode == "Nach Prio" and not overview["prio_counts"].empty:
                    chart_style = st.radio(
                        "Zeitaufwand darstellen als",
                        ["Zweite Achse", "Wert im Balken"],
                        horizontal=True,
                        key="phase_overview_style",
                    )
                    chart_html = _build_overview_svg(
                        overview["prio_counts"],
                        "prio_stage",
                        "messstellen",
                        "aufwand_plus_puffer_min",
                        chart_style,
                        "Messstellen und Aufwand nach Prio",
                        "Aufwand (h)",
                    )
                    if chart_html:
                        st.components.v1.html(chart_html, height=430, scrolling=False)
                    st.dataframe(overview["prio_counts"], use_container_width=True, height=180, column_config=build_column_config(overview["prio_counts"], allow_manual_edit=False))
                elif view_mode == "Nach Phase" and not overview["phase_counts"].empty:
                    chart_style = st.radio(
                        "Zeitaufwand darstellen als",
                        ["Zweite Achse", "Wert im Balken"],
                        horizontal=True,
                        key="phase_overview_style",
                    )
                    chart_html = _build_overview_svg(
                        overview["phase_counts"],
                        "phase",
                        "messstellen",
                        "aufwand_plus_puffer_min",
                        chart_style,
                        "Messstellen und Aufwand nach Phase",
                        "Aufwand (h)",
                    )
                    if chart_html:
                        st.components.v1.html(chart_html, height=430, scrolling=False)
                    st.dataframe(overview["phase_counts"], use_container_width=True, height=180, column_config=build_column_config(overview["phase_counts"], allow_manual_edit=False))
                elif view_mode == "Phase x Prio" and not overview["phase_prio_matrix"].empty:
                    chart_df = overview["phase_prio_matrix"].set_index("phase")
                    st.bar_chart(chart_df, height=360)
                    st.dataframe(overview["phase_prio_matrix"], use_container_width=True, height=220, column_config=build_column_config(overview["phase_prio_matrix"], allow_manual_edit=False))
            else:
                st.info("Kein grafischer Überblick verfügbar.")

            raw_export = df_in.copy()
            raw_export = raw_export[[c for c in raw_export.columns if c in raw_export.columns]]
            raw_export_preview_cols = [
                c
                for c in [
                    "Asset ID",
                    "decision_status",
                    "prio_stage",
                    "prio_substage",
                    "relabel_phase",
                    "recommended_window",
                    "unterbruch_erforderlich",
                    "decision_reason",
                    "time_class",
                    "access_class",
                    "ms_legacy_class",
                    "qc_scope_status",
                ]
                if c in raw_export.columns
            ]
            export_meta = {
                "generated_at": datetime.now().isoformat(timespec="seconds"),
                "source_csv_path": source_path or st.session_state.get("last_selected_csv_path", ""),
                "source_csv_name": Path(source_path).name if source_path else Path(str(st.session_state.get("last_selected_csv_path", ""))).name,
                "decision_tree_path": st.session_state.get("active_criteria_path", ""),
                "decision_tree_name": str(st.session_state.get("active_criteria_set", {}).get("name", "")),
                "decision_tree_version": str(st.session_state.get("active_criteria_set", {}).get("version", "")),
                "decision_tree": summarize_decision_tree(st.session_state.get("active_criteria_set", {})),
                "capacity": {
                    "min_normal": min_normal,
                    "min_medium": min_medium,
                    "buffer_pct": buffer_pct,
                    "persons": persons,
                    "hours_day": hours_day,
                },
                "summary": {
                    "phases": int(grp["relabel_phase"].nunique()),
                    "rows": int(grp["anzahl_messstellen"].sum()),
                    "total_personentage": float(grp["personentage"].sum()),
                    "total_wochenbedarf": float(grp["wochenbedarf"].sum()),
                },
                "raw_classified_preview": json_safe(
                    raw_export[raw_export_preview_cols].head(50).to_dict(orient="records")
                    if raw_export_preview_cols
                    else raw_export.head(50).to_dict(orient="records")
                ),
                "phase_plan_preview": json_safe(grp.head(20).to_dict(orient="records")),
            }
            export_csv = grp.to_csv(index=False).encode("utf-8-sig")
            raw_export_csv = raw_export.to_csv(index=False).encode("utf-8-sig")
            export_xlsx = build_phase_plan_excel_export(grp, raw_export, export_meta)
            export_zip = build_named_export_zip(
                {
                    "phase_plan_aggregated.csv": export_csv,
                    "phase_plan_raw_classified.csv": raw_export_csv,
                    "phase_plan_metadata.json": json.dumps(export_meta, ensure_ascii=False, indent=2),
                }
            )
            cexp1, cexp2, cexp3 = st.columns(3)
            cexp1.download_button(
                "Phasenplan als CSV exportieren",
                export_csv,
                file_name=f"phase_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="phase_plan_export_csv",
            )
            cexp2.download_button(
                "Phasenplan als Excel exportieren",
                export_xlsx,
                file_name=f"phase_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="phase_plan_export_xlsx",
            )
            cexp3.download_button(
                "Phasenplan + Metadaten (ZIP)",
                export_zip,
                file_name=f"phase_plan_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                key="phase_plan_export_zip",
            )
    with help_col:
        render_context_help("phase_plan")

def render_criteria_comparison(df: pd.DataFrame) -> None:
    st.markdown("### Entscheidungsbaum-Vergleich")
    criteria_files = list_criteria_sets()
    if len(criteria_files) < 2:
        st.info("Für den Vergleich werden mindestens 2 Entscheidungsbäume benötigt.")
        return

    options = {p.name: str(p) for p in criteria_files}
    names = list(options.keys())

    c1, c2 = st.columns(2)
    left_name = c1.selectbox("Version A", names, index=0, key="cmp_left")
    right_name = c2.selectbox("Version B", names, index=1 if len(names) > 1 else 0, key="cmp_right")

    left = load_criteria(options[left_name])
    right = load_criteria(options[right_name])
    ref_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}

    left_in, left_ex = apply_criteria_rules(df.copy(), left, ref_lists)
    right_in, right_ex = apply_criteria_rules(df.copy(), right, ref_lists)

    def metrics(in_df: pd.DataFrame, ex_df: pd.DataFrame) -> Dict[str, int]:
        return {
            "in_prio": int(len(in_df)),
            "excluded": int(len(ex_df)),
            "P1": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P1").sum()),
            "P2": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P2").sum()),
            "P2A": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P2A").sum()),
            "P3": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P3").sum()),
            "P4": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P4").sum()),
            "P5": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P5").sum()),
            "P6": int((in_df.get("prio_stage", pd.Series(dtype='string')) == "P6").sum()),
            "unterbruch": int((in_df.get("unterbruch_erforderlich", pd.Series(dtype='bool')) == True).sum()),
        }

    m1 = metrics(left_in, left_ex)
    m2 = metrics(right_in, right_ex)

    rows = []
    for k in ["in_prio", "excluded", "P1", "P2", "P2A", "P3", "P4", "P5", "P6", "unterbruch"]:
        rows.append({"kennzahl": k, "A": m1[k], "B": m2[k], "delta_B_minus_A": m2[k] - m1[k]})
    cmp_df = pd.DataFrame(rows)
    st.dataframe(cmp_df, use_container_width=True, height=260, column_config=build_column_config(cmp_df, allow_manual_edit=False))

    st.markdown("#### Unterschiedsliste (Zeilenstatus)")
    id_col = "Asset ID" if "Asset ID" in df.columns else None
    if not id_col:
        st.info("Keine Spalte 'Asset ID' gefunden. Statusvergleich pro Zeile nicht verfügbar.")
        return

    a_status = left_in[[id_col, "decision_status", "prio_stage", "prio_substage", "unterbruch_erforderlich"]].copy()
    b_status = right_in[[id_col, "decision_status", "prio_stage", "prio_substage", "unterbruch_erforderlich"]].copy()

    a_status = a_status.rename(columns={
        "decision_status": "status_A",
        "prio_stage": "prio_A",
        "prio_substage": "subprio_A",
        "unterbruch_erforderlich": "unterbruch_A",
    })
    b_status = b_status.rename(columns={
        "decision_status": "status_B",
        "prio_stage": "prio_B",
        "prio_substage": "subprio_B",
        "unterbruch_erforderlich": "unterbruch_B",
    })

    # Add excluded entries as status rows.
    if not left_ex.empty:
        exa = left_ex[[id_col]].copy()
        exa["status_A"] = "excluded"
        exa["prio_A"] = "excluded"
        exa["subprio_A"] = ""
        exa["unterbruch_A"] = left_ex.get("unterbruch_erforderlich", False).values
        a_status = pd.concat([a_status, exa], ignore_index=True)
    if not right_ex.empty:
        exb = right_ex[[id_col]].copy()
        exb["status_B"] = "excluded"
        exb["prio_B"] = "excluded"
        exb["subprio_B"] = ""
        exb["unterbruch_B"] = right_ex.get("unterbruch_erforderlich", False).values
        b_status = pd.concat([b_status, exb], ignore_index=True)

    merged = a_status.merge(b_status, on=id_col, how="outer").fillna("")
    changed = merged[
        (merged["status_A"] != merged["status_B"]) |
        (merged["prio_A"] != merged["prio_B"]) |
        (merged["subprio_A"] != merged["subprio_B"]) |
        (merged["unterbruch_A"].astype(str) != merged["unterbruch_B"].astype(str))
    ].copy()

    st.write(f"{len(changed)} geänderte Zeilen zwischen A und B")
    show_cols = [id_col, "status_A", "status_B", "prio_A", "prio_B", "subprio_A", "subprio_B", "unterbruch_A", "unterbruch_B"]
    st.dataframe(changed[show_cols], use_container_width=True, height=360, column_config=build_column_config(changed[show_cols], allow_manual_edit=False))



def render_help_docs() -> None:
    st.markdown("### Hilfe & Anleitung")
    doku_dir = BASE_DIR / "doku"
    guide = doku_dir / "ANLEITUNG.md"
    click_guide = doku_dir / "KLICKANLEITUNG.md"

    t1, t2 = st.tabs(["ANLEITUNG", "KLICKANLEITUNG"])
    with t1:
        if guide.exists():
            st.markdown(guide.read_text(encoding="utf-8"))
        else:
            st.warning(f"Nicht gefunden: {guide}")
    with t2:
        if click_guide.exists():
            st.markdown(click_guide.read_text(encoding="utf-8"))
        else:
            st.warning(f"Nicht gefunden: {click_guide}")

def main() -> None:
    apply_custom_style()
    restore_last_session()
    sanitize_filter_stores()
    st.title("KAU Relabeling Priorisierung")

    # Initialize default criteria set and persisted QC reference list.
    if "criteria_initialized" not in st.session_state:
        default_path = ensure_default_criteria_file()
        restored_criteria_path = (
            str(st.session_state["active_criteria_path"]).strip()
            if "active_criteria_path" in st.session_state
            else None
        )
        if restored_criteria_path is None:
            st.session_state["active_criteria_path"] = str(default_path)
            st.session_state["active_criteria_set"] = load_criteria(str(default_path))
        elif restored_criteria_path == "":
            st.session_state["active_criteria_path"] = ""
            st.session_state["active_criteria_set"] = {"name": "none", "version": "", "notes": "", "rules": []}
        elif restored_criteria_path and Path(restored_criteria_path).exists():
            st.session_state["active_criteria_path"] = restored_criteria_path
            st.session_state["active_criteria_set"] = load_criteria(restored_criteria_path)
        else:
            st.session_state["active_criteria_path"] = str(default_path)
            st.session_state["active_criteria_set"] = load_criteria(str(default_path))
        set_active_decision_tree_config(st.session_state.get("active_criteria_set", {}).get("config"))
        qc_csv = REF_DIR / "qc_self_labeled_assets.csv"
        if qc_csv.exists():
            try:
                qc_assets = pd.read_csv(qc_csv)["Asset ID"].dropna().astype(str).str.strip().tolist()
                st.session_state["qc_green_assets"] = sorted(set(qc_assets))
                st.session_state["qc_green_source"] = str(qc_csv)
            except Exception:
                pass
        ov_csv = REF_DIR / "manual_overrides.csv"
        if ov_csv.exists():
            try:
                ov_raw = pd.read_csv(ov_csv)
                st.session_state["manual_overrides_df"] = normalize_override_rows(ov_raw)
                st.session_state["manual_overrides_source"] = str(ov_csv)
            except Exception:
                pass
        st.session_state["criteria_initialized"] = True


    st.sidebar.header("Arbeitsmodus")
    sidebar_mode = st.sidebar.radio(
        "Sidebar anzeigen für",
        ["Filter bearbeiten", "Daten & Referenzen", "Entscheidungsbaum & Szenarien", "Kapazität", "Hilfe"],
        key="sidebar_mode",
    )

    show_filters = sidebar_mode == "Filter bearbeiten"
    show_data = sidebar_mode == "Daten & Referenzen"
    show_criteria = sidebar_mode == "Entscheidungsbaum & Szenarien"
    show_capacity = sidebar_mode == "Kapazität"
    show_help = sidebar_mode == "Hilfe"

    csv_files = list_local_csvs()
    selected_path = None
    csv_options = {p.name: str(p) for p in csv_files}
    csv_option_names = list(csv_options.keys())
    if csv_option_names:
        last_csv_name = Path(str(st.session_state.get("last_selected_csv_path", ""))).name
        default_csv_name = st.session_state.get("selected_csv_name") or last_csv_name
        if default_csv_name not in csv_options:
            default_csv_name = csv_option_names[0]
        if "selected_csv_name" not in st.session_state or st.session_state["selected_csv_name"] not in csv_options:
            st.session_state["selected_csv_name"] = default_csv_name
        selected_path = csv_options[st.session_state["selected_csv_name"]]
        st.session_state["last_selected_csv_path"] = selected_path

    if show_help:
        st.sidebar.header("Hilfe")
        st.sidebar.caption("Anleitungen findest du im Tab 'Hilfe & Anleitung'.")
        st.sidebar.caption("Die rechte Hilfe-Spalte in komplexen Tabs erklärt die wichtigsten Begriffe direkt im Kontext.")

    if show_data:
        st.sidebar.header("Datenquelle")
        with st.sidebar.expander("CSV Auswahl / Upload", expanded=True):
            if csv_option_names:
                selected_name = st.selectbox(
                    "Aktuelle CSV",
                    csv_option_names,
                    index=csv_option_names.index(st.session_state["selected_csv_name"]),
                    key="selected_csv_name",
                )
                selected_path = csv_options[selected_name]
                st.session_state["last_selected_csv_path"] = selected_path
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
                    st.session_state["last_selected_csv_path"] = selected_path

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

        with st.sidebar.expander("Manuelle Overrides (CSV)", expanded=False):
            st.caption("Kleine Ausnahmeliste für Einzelfälle. Spalten: Asset ID, override_action, override_prio, override_substage, override_reason, override_shutdown")
            ov_file = st.file_uploader("Override-CSV hochladen", type=["csv"], key="override_csv_upload")
            if ov_file is not None and st.button("Override-Liste einlesen", key="override_load_btn"):
                ov_path = REF_DIR / "manual_overrides.csv"
                ov_path.write_bytes(ov_file.getvalue())
                ov_raw = pd.read_csv(ov_path)
                ov_norm = normalize_override_rows(ov_raw)
                st.session_state["manual_overrides_df"] = ov_norm
                st.session_state["manual_overrides_source"] = str(ov_path)
                st.success(f"Override-Liste geladen: {len(ov_norm)} Assets")

            if "manual_overrides_df" not in st.session_state or st.session_state.get("manual_overrides_df") is None:
                st.session_state["manual_overrides_df"] = empty_override_table()

            st.caption("Direkt hier pflegen: Zeilen hinzufügen/ändern und speichern.")
            edit_df = st.data_editor(
                st.session_state.get("manual_overrides_df", empty_override_table()),
                key="override_editor",
                num_rows="dynamic",
                use_container_width=True,
                column_config={
                    "override_action": st.column_config.SelectboxColumn(options=["assign_prio", "exclude_from_prio"]),
                    "override_prio": st.column_config.SelectboxColumn(options=["", "P1", "P2", "P2A", "P3", "P4", "P5", "P6"]),
                    "override_shutdown": st.column_config.CheckboxColumn(),
                },
            )
            c_ov1, c_ov2 = st.columns(2)
            if c_ov1.button("Override-Liste übernehmen", key="override_apply_btn"):
                ov_norm = normalize_override_rows(edit_df)
                st.session_state["manual_overrides_df"] = ov_norm
                st.success(f"Override-Liste übernommen: {len(ov_norm)} Assets")
            if c_ov2.button("Override-Liste speichern", key="override_save_btn"):
                ov_path = REF_DIR / "manual_overrides.csv"
                ov_norm = normalize_override_rows(edit_df)
                ov_norm.to_csv(ov_path, index=False)
                st.session_state["manual_overrides_df"] = ov_norm
                st.session_state["manual_overrides_source"] = str(ov_path)
                st.success(f"Gespeichert: {ov_path}")

            ov_df = st.session_state.get("manual_overrides_df", empty_override_table())
            st.caption(f"Aktive Override-Liste: {len(ov_df)} Assets")
            if st.session_state.get("manual_overrides_source"):
                st.caption(f"Quelle: {st.session_state['manual_overrides_source']}")

        with st.sidebar.expander("Arbeitsstand", expanded=False):
            saved_at = st.session_state.get("_last_session_saved_at", "")
            if saved_at:
                st.caption(f"Letzter automatisch geladener Stand: {saved_at}")
            else:
                st.caption("Noch kein gespeicherter Arbeitsstand vorhanden.")
            if st.session_state.get("_last_session_restore_error"):
                st.warning(f"Wiederherstellung nicht möglich: {st.session_state['_last_session_restore_error']}")
            if st.session_state.get("_last_session_save_error"):
                st.warning(f"Speichern zuletzt nicht möglich: {st.session_state['_last_session_save_error']}")
            c_state_1, c_state_2 = st.columns(2)
            if c_state_1.button("Jetzt speichern", key="last_session_save_now"):
                save_last_session(selected_path)
                st.session_state["_last_session_saved_at"] = datetime.now().isoformat(timespec="seconds")
                st.success("Arbeitsstand gespeichert.")
            if c_state_2.button("Zurücksetzen", key="last_session_reset"):
                reset_last_session()
                st.session_state["_last_session_saved_at"] = ""
                st.session_state["_skip_last_session_autosave"] = True
                st.success("Gespeicherter Arbeitsstand gelöscht. Aktuelle Ansicht bleibt bis zum Neustart unverändert.")
    if "selected_path" not in locals() or not selected_path:
        st.info("Bitte CSV im Arbeitsmodus 'Daten & Referenzen' auswählen oder hochladen.")
        st.stop()

    if "pending_active_prefix" in st.session_state:
        st.session_state["active_prefix"] = st.session_state.pop("pending_active_prefix")
    if "active_prefix" not in st.session_state:
        st.session_state["active_prefix"] = "view_a"
    if show_filters:
        st.sidebar.header("Ansicht")
        active_prefix = st.sidebar.selectbox(
            "Filter bearbeiten für",
            options=["view_a", "view_b", "view_c"],
            format_func=lambda x: {"view_a": "Filteransicht A", "view_b": "Filteransicht B", "view_c": "Filteransicht C"}[x],
            key="active_prefix",
        )
    else:
        active_prefix = st.session_state["active_prefix"]
        st.sidebar.caption(f"Aktive Filteransicht: { {'view_a': 'A', 'view_b': 'B', 'view_c': 'C'}.get(active_prefix, active_prefix) }")

    if "golive_reference_date" not in st.session_state:
        st.session_state["golive_reference_date"] = date(2026, 8, 10)
    if show_criteria:
        st.sidebar.header("Entscheidungsbaum")
        st.sidebar.date_input(
            "Go-Live Referenzdatum",
            value=st.session_state["golive_reference_date"],
            key="golive_reference_date",
            help="Wird für die abgeleitete Kennzahl months_from_golive verwendet.",
        )
        criteria_files = list_criteria_sets()
        if criteria_files:
            options = {"(kein Entscheidungsbaum)": ""}
            options.update({p.name: str(p) for p in criteria_files})

            current_path = st.session_state.get("active_criteria_path", str(criteria_files[0]))
            current_name = Path(current_path).name if current_path else "(kein Entscheidungsbaum)"
            if current_name not in options:
                current_name = list(options.keys())[0]

            selected_criteria_name = st.sidebar.selectbox(
                "Aktiver Entscheidungsbaum",
                list(options.keys()),
                index=list(options.keys()).index(current_name),
            )
            selected_criteria_path = options[selected_criteria_name]

            if selected_criteria_path == "":
                if st.session_state.get("active_criteria_path") != "":
                    st.session_state["active_criteria_path"] = ""
                    st.session_state["active_criteria_set"] = {"name": "none", "version": "", "notes": "", "rules": []}
                    set_active_decision_tree_config(None)
                    st.rerun()
            elif selected_criteria_path != st.session_state.get("active_criteria_path"):
                st.session_state["active_criteria_path"] = selected_criteria_path
                st.session_state["active_criteria_set"] = load_criteria(selected_criteria_path)
                set_active_decision_tree_config(st.session_state["active_criteria_set"].get("config"))
                st.rerun()

    if show_capacity:
        st.sidebar.header("Kapazität")
        st.sidebar.number_input("Minuten/Messstelle normal", min_value=1.0, value=float(st.session_state.get("cap_min_normal", 7)), key="cap_min_normal")
        st.sidebar.number_input("Minuten/Messstelle mittel", min_value=1.0, value=float(st.session_state.get("cap_min_medium", 12)), key="cap_min_medium")
        st.sidebar.number_input("Puffer %", min_value=0.0, value=float(st.session_state.get("cap_buffer_pct", 30)), key="cap_buffer_pct")
        st.sidebar.number_input("Personenanzahl", min_value=1.0, value=float(st.session_state.get("cap_persons", 2)), key="cap_persons")
        st.sidebar.number_input("Stunden pro Person/Tag", min_value=1.0, value=float(st.session_state.get("cap_hours_day", 7)), key="cap_hours_day")

    df_raw = load_csv_from_path(selected_path)

    if "datalogger_prefilter_on" not in st.session_state:
        st.session_state["datalogger_prefilter_on"] = False
    raw_text_cols = get_text_columns(df_raw)
    datalogger_candidates = [c for c in raw_text_cols if any(k in c.lower() for k in ["beschreibung", "description", "desc", "messstellen"])]
    if not datalogger_candidates:
        datalogger_candidates = raw_text_cols
    with st.sidebar.expander("Datenlogger-Vorfilter", expanded=False):
        if datalogger_candidates:
            default_description = find_description_column(df_raw)
            if default_description not in datalogger_candidates:
                default_description = datalogger_candidates[0]
            description_idx = datalogger_candidates.index(default_description) if default_description in datalogger_candidates else 0
            datalogger_description_col = st.selectbox(
                "Messstellenbeschreibung-Spalte",
                datalogger_candidates,
                index=description_idx,
                key="datalogger_prefilter_desc_col",
            )
            st.checkbox(
                "Datenlogger vorab aus dem Phasenplan entfernen",
                value=bool(st.session_state.get("datalogger_prefilter_on", False)),
                key="datalogger_prefilter_on",
            )
            datalogger_count_raw, _ = count_datalogger_rows(df_raw, description_col=datalogger_description_col)
            st.caption(f"Datenlogger in Originaldatei: {datalogger_count_raw}")
            st.caption("Wirkung: entfernt die Zeilen vor Entscheidungsbaum und Phasenplan.")
        else:
            st.caption("Keine geeignete Textspalte für die Messstellenbeschreibung gefunden.")

    df, datalogger_removed_count, datalogger_description_col = apply_datalogger_prefilter(
        df_raw,
        bool(st.session_state.get("datalogger_prefilter_on", False)),
        st.session_state.get("datalogger_prefilter_desc_col"),
    )
    st.session_state["datalogger_prefilter_removed_count"] = datalogger_removed_count
    st.session_state["datalogger_prefilter_description_col"] = datalogger_description_col or ""
    df = derive_decision_tree_columns(df)
    if show_criteria:
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

    ref_lists = {"qc_self_labeled_assets": set(st.session_state.get("qc_green_assets", []))}
    criteria_all_in, criteria_all_ex = apply_criteria_rules(df.copy(), st.session_state.get("active_criteria_set", {}), ref_lists)

    tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10, tab11, tab12 = st.tabs(["Filteransicht A", "Filteransicht B", "Filteransicht C", "CSV-Check", "Standort-Analyse", "Entscheidungsbaum-Konfig", "Vergleich", "Entscheidungsbaum", "Phasenplan", "Prio-Matrix", "Ausschlüsse", "Hilfe & Anleitung"])
    sidebar_filter_prefix = active_prefix if show_filters else "__none__"
    with tab1:
        render_view(df, "Filteransicht A", "view_a", sidebar_filter_prefix)
    with tab2:
        render_view(df, "Filteransicht B", "view_b", sidebar_filter_prefix)
    with tab3:
        render_view(df, "Filteransicht C", "view_c", sidebar_filter_prefix)
    with tab4:
        render_csv_check(df)
    with tab5:
        render_standort_analyse(df)
    with tab6:
        render_criteria_editor(df)
    with tab7:
        render_criteria_comparison(df)
    with tab8:
        render_decision_tree(df)
    with tab9:
        render_phase_plan(criteria_all_in, source_path=selected_path)
    with tab10:
        render_prio_matrix(criteria_all_in)
    with tab11:
        render_exclusions(criteria_all_ex)
    with tab12:
        render_help_docs()

    if st.session_state.pop("_skip_last_session_autosave", False):
        pass
    else:
        save_last_session(selected_path)


if __name__ == "__main__":
    main()
