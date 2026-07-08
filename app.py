import html
import json
import math
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from docx import Document
from docx.shared import Pt


APP_NAME = "Paper Parallel Reader"
APP_VERSION = "1.5.4-progress-selection-polish"

STATUS_OPTIONS = ["未確認", "確認済み", "要修正", "修正済み"]
GLOSSARY_CATEGORIES = ["専門用語", "理論概念", "方法", "固有名詞", "その他"]
WORD_EXPORT_OPTIONS = [
    "2列表（元文｜訳文）",
    "縦並び（元文→訳文）",
    "訳文のみ",
    "確認用（ID・状態・メモつき）",
]

DEFAULT_VIEWER_HEIGHT = 560

st.set_page_config(page_title=APP_NAME, layout="wide")


# ============================================================
# デザイン
# ============================================================
def inject_style():
    st.markdown(
        """
<style>
:root {
  --main-blue: #334f63;
  --deep-blue: #17212b;
  --soft-blue: #edf2f4;
  --main-green: #2f7d68;
  --soft-green: #edf7f2;
  --paper: #faf8f3;
  --ink: #1f2933;
  --muted: #6b7280;
  --line: #e3ded4;
  --ppr-unchecked: #d7d2c8;
  --ppr-needsfix: #c68245;
  --ppr-confirmed: #4f6f88;
  --ppr-revised: #2f7d68;
}

.block-container {
  padding-top: 1.2rem;
  padding-bottom: 3rem;
}

.ppr-hero {
  background: linear-gradient(135deg, #18222d 0%, #2f4658 58%, #48685f 100%);
  color: white;
  border-radius: 22px;
  padding: 22px 26px;
  box-shadow: 0 14px 34px rgba(31, 41, 51, .22);
  margin-bottom: 18px;
}
.ppr-hero h1 {
  margin: 0 0 6px 0;
  font-size: 2.05rem;
  letter-spacing: .01em;
}
.ppr-hero p {
  margin: 0;
  color: rgba(255,255,255,.88);
  font-size: .98rem;
}
.ppr-badge {
  display:inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  background: rgba(255,255,255,.14);
  border: 1px solid rgba(255,255,255,.24);
  margin-right: 6px;
  margin-top: 10px;
  font-size: .82rem;
}
.ppr-card {
  background: white;
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 14px 16px;
  box-shadow: 0 10px 24px rgba(15,23,42,.05);
}
.ppr-subtle {
  color: var(--muted);
  font-size: .92rem;
}
.ppr-chip {
  display:inline-block;
  padding: 4px 10px;
  border-radius: 999px;
  border: 1px solid var(--line);
  background: var(--paper);
  margin: 2px 4px 2px 0;
  font-size: .82rem;
}
.ppr-chip-blue { background: var(--soft-blue); border-color:#cfd8dc; color:#334f63; }
.ppr-chip-green { background: var(--soft-green); border-color:#c7ddd3; color:#2f7d68; }
.ppr-chip-yellow { background:#f7efe1; border-color:#e5c79c; color:#8a5a2b; }
.ppr-chip-red { background:#f4e5e0; border-color:#e5b4a7; color:#8f3f2d; }

[data-testid="stMetricValue"] { color: var(--deep-blue); }

.stTabs [data-baseweb="tab-list"] {
  gap: 8px;
  border-bottom: 1px solid var(--line);
}
.stTabs [data-baseweb="tab"] {
  border-radius: 12px 12px 0 0;
  padding: 10px 14px;
  font-weight: 700;
}
.stTabs [aria-selected="true"] {
  background: var(--soft-blue);
  color: var(--main-blue);
}

div.stButton > button:first-child,
div.stDownloadButton > button:first-child {
  border-radius: 12px;
  border: 1px solid #bcd1ee;
  font-weight: 700;
}

section[data-testid="stSidebar"] {
  background: linear-gradient(180deg, #fbfaf7 0%, #f3f0e8 100%);
}

.ppr-progress-card {
  background: rgba(255,255,255,.88);
  border: 1px solid var(--line);
  border-radius: 18px;
  padding: 14px 14px 16px;
  box-shadow: 0 10px 24px rgba(31,41,51,.06);
}
.ppr-progress-head {
  display: flex;
  justify-content: space-between;
  align-items: baseline;
  gap: 8px;
  margin-bottom: 12px;
}
.ppr-progress-title {
  font-weight: 900;
  color: var(--deep-blue);
  letter-spacing: .02em;
}
.ppr-progress-percent {
  font-weight: 900;
  color: var(--ppr-revised);
  font-size: 1.12rem;
  font-variant-numeric: tabular-nums;
}
.ppr-progress-body {
  display: flex;
  justify-content: center;
  align-items: center;
}
.ppr-progress-stack {
  width: 112px;
  height: 104px;
  position: relative;
  display: flex;
  flex-direction: column;
  border-radius: 18px;
  background: #eee9df;
  box-shadow: inset 0 0 0 1px rgba(255,255,255,.58), 0 8px 18px rgba(31,41,51,.10);
  overflow: visible;
}
.ppr-progress-segment {
  position: relative;
  width: 100%;
  min-height: 4px;
  cursor: help;
  transition: filter .12s ease, transform .12s ease;
}
.ppr-progress-segment:hover {
  filter: brightness(.94) saturate(1.08);
  z-index: 15;
}
.ppr-progress-segment:first-child {
  border-radius: 18px 18px 0 0;
}
.ppr-progress-segment:last-child {
  border-radius: 0 0 18px 18px;
}
.ppr-progress-segment:only-child {
  border-radius: 18px;
}
.ppr-progress-unchecked { background: var(--ppr-unchecked); }
.ppr-progress-confirmed { background: var(--ppr-confirmed); }
.ppr-progress-needsfix { background: var(--ppr-needsfix); }
.ppr-progress-revised { background: var(--ppr-revised); }
.ppr-progress-tip {
  display: none;
  position: absolute;
  left: 50%;
  top: 50%;
  transform: translate(-50%, -50%);
  width: max-content;
  min-width: 154px;
  max-width: 190px;
  padding: 9px 11px;
  border-radius: 13px;
  background: rgba(24,34,45,.96);
  color: #fffdf7;
  font-size: 12.5px;
  line-height: 1.45;
  z-index: 50;
  box-shadow: 0 12px 26px rgba(31,41,51,.24);
  pointer-events: none;
  white-space: normal;
}
.ppr-progress-segment:hover .ppr-progress-tip {
  display: block;
}
.ppr-progress-tip-label {
  display: block;
  font-weight: 900;
  color: #f2eadb;
  margin-bottom: 2px;
}
.ppr-progress-tip-detail {
  display: block;
  color: rgba(255,253,247,.86);
  font-variant-numeric: tabular-nums;
}
</style>
""",
        unsafe_allow_html=True,
    )


# ============================================================
# 基本ユーティリティ
# ============================================================
def init_state():
    defaults = {
        "rows": [],
        "glossary": [],
        "project_title": "paper_parallel_project",
        "unit": "段落",
        "selected_id": None,
        "active_page": "一覧・選択",
        "page_size": 20,
        "context_radius": 2,
        "show_settings": True,
        "show_json_loader": True,
        "show_text_adder": True,
        "last_filtered_ids": [],
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value


def now_string():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def safe_filename(text, fallback="paper_parallel"):
    text = str(text or fallback).strip() or fallback
    text = re.sub(r"[\\/:*?\"<>|\s]+", "_", text)
    return text[:80]


def truncate_text(text, max_chars=120):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    return text if len(text) <= max_chars else text[:max_chars] + "…"


def estimate_english_words(text):
    return len(re.findall(r"[A-Za-z]+(?:[-'][A-Za-z]+)?", str(text or "")))


def get_final_translation(row):
    edited = str(row.get("修正訳", "")).strip()
    return edited if edited else str(row.get("原訳", "")).strip()


def normalize_status(status):
    status = str(status or "未確認")
    if status in STATUS_OPTIONS:
        return status
    if status == "手動":
        return "修正済み"
    return "未確認"


def clean_phrase(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def lower_key(text):
    return clean_phrase(text).lower()


def next_item_id(items):
    ids = []
    for item in items:
        try:
            ids.append(int(item.get("id", 0)))
        except Exception:
            pass
    return max(ids or [0]) + 1


# ============================================================
# 分割関数
# ============================================================
def split_paragraphs(text):
    text = str(text or "").strip()
    if not text:
        return []
    paragraphs = re.split(r"\n\s*\n+", text)
    return [re.sub(r"\s+", " ", p.strip()) for p in paragraphs if p.strip()]


def split_english_sentences(text):
    text = re.sub(r"\s+", " ", str(text or "").strip())
    if not text:
        return []
    sentences = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9\"'(\[])", text)
    return [s.strip() for s in sentences if s.strip()]


def split_japanese_sentences(text):
    text = re.sub(r"\s+", " ", str(text or "").strip())
    if not text:
        return []
    sentences = re.split(r"(?<=[。！？!?])\s*", text)
    return [s.strip() for s in sentences if s.strip()]


def segment_text(text, language, unit):
    if unit == "段落":
        return split_paragraphs(text)
    return split_english_sentences(text) if language == "英語" else split_japanese_sentences(text)


# ============================================================
# 行データ
# ============================================================
def make_rows_from_segments(english_segments, japanese_segments, section_title, start_id=1):
    section_title = str(section_title or "Section").strip() or "Section"
    row_count = max(len(english_segments), len(japanese_segments))
    rows = []
    for i in range(row_count):
        en = english_segments[i] if i < len(english_segments) else ""
        ja = japanese_segments[i] if i < len(japanese_segments) else ""
        rows.append(
            {
                "ID": start_id + i,
                "章": section_title,
                "章内ID": i + 1,
                "英文": en,
                "原訳": ja,
                "修正訳": ja,
                "状態": "未確認",
                "メモ": "",
                "更新日時": "",
            }
        )
    return rows


def normalize_row(row, fallback_id, fallback_section="Imported"):
    if "ID" in row and "英文" in row:
        original = str(row.get("原訳", row.get("和訳", row.get("訳文", ""))))
        edited = str(row.get("修正訳", row.get("和訳（編集用）", row.get("訳文", original))))
        return {
            "ID": int(row.get("ID") or fallback_id),
            "章": str(row.get("章", fallback_section) or fallback_section),
            "章内ID": int(row.get("章内ID") or fallback_id),
            "英文": str(row.get("英文", "")),
            "原訳": original,
            "修正訳": edited,
            "状態": normalize_status(row.get("状態", row.get("修正状態", "未確認"))),
            "メモ": str(row.get("メモ", "")),
            "更新日時": str(row.get("更新日時", "")),
        }

    original = str(row.get("和訳", row.get("訳文", "")))
    edited = str(row.get("和訳（編集用）", original))
    status = "修正済み" if row.get("修正状態") == "手動" else "未確認"
    return {
        "ID": int(row.get("英語ID") or fallback_id),
        "章": fallback_section,
        "章内ID": int(row.get("英語ID") or fallback_id),
        "英文": str(row.get("英文", "")),
        "原訳": original,
        "修正訳": edited,
        "状態": status,
        "メモ": "",
        "更新日時": "",
    }


def get_row_index_by_id(rows, row_id):
    for index, row in enumerate(rows):
        try:
            if int(row.get("ID", -1)) == int(row_id):
                return index
        except Exception:
            continue
    return None


def get_row_by_id(rows, row_id):
    index = get_row_index_by_id(rows, row_id)
    return rows[index] if index is not None else None


def renumber_rows(rows):
    section_counts = {}
    for index, row in enumerate(rows, start=1):
        section = str(row.get("章", "Section") or "Section")
        section_counts[section] = section_counts.get(section, 0) + 1
        row["ID"] = index
        row["章内ID"] = section_counts[section]
    return rows


# ============================================================
# 用語辞典
# ============================================================
def glossary_unique_key(item):
    return (
        lower_key(item.get("english", "")),
        clean_phrase(item.get("japanese", "")),
        str(item.get("category", "")),
    )


def cleanup_glossary():
    glossary = []
    seen = set()
    for item in st.session_state.get("glossary", []):
        english = clean_phrase(item.get("english", item.get("英語", "")))
        japanese = clean_phrase(item.get("japanese", item.get("日本語", "")))
        if not english and not japanese:
            continue
        fixed = {
            "id": int(item.get("id", len(glossary) + 1) or len(glossary) + 1),
            "english": english,
            "japanese": japanese,
            "category": str(item.get("category", item.get("カテゴリ", "専門用語")) or "専門用語"),
            "note": str(item.get("note", item.get("メモ", ""))),
            "created_at": str(item.get("created_at", "")),
        }
        key = glossary_unique_key(fixed)
        if key in seen:
            continue
        seen.add(key)
        glossary.append(fixed)

    for i, item in enumerate(glossary, start=1):
        item["id"] = i
    st.session_state["glossary"] = glossary


def add_unique_glossary(english, japanese, category, note):
    item = {
        "id": next_item_id(st.session_state["glossary"]),
        "english": clean_phrase(english),
        "japanese": clean_phrase(japanese),
        "category": category,
        "note": str(note or "").strip(),
        "created_at": now_string(),
    }
    if not item["english"] and not item["japanese"]:
        return False, "英語用語または日本語訳を入力してください。"

    existing_keys = {glossary_unique_key(g) for g in st.session_state["glossary"]}
    if glossary_unique_key(item) in existing_keys:
        return False, "同じ用語はすでに登録されています。"

    st.session_state["glossary"].append(item)
    cleanup_glossary()
    return True, "用語辞典に追加しました。"


def has_glossary_hit(row):
    en = str(row.get("英文", ""))
    for item in st.session_state.get("glossary", []):
        english = clean_phrase(item.get("english", ""))
        if english and re.search(make_term_pattern(english), en, flags=re.IGNORECASE):
            return True
    return False


def make_term_pattern(term):
    escaped = re.escape(clean_phrase(term))
    if not escaped:
        return r"a^"
    left = r"(?<![A-Za-z0-9_-])" if re.match(r"^[A-Za-z0-9]", term) else ""
    right = r"(?![A-Za-z0-9_-])" if re.search(r"[A-Za-z0-9]$", term) else ""
    return f"{left}{escaped}{right}"


# ============================================================
# JSON読み込み・保存
# ============================================================
def load_rows_from_json(uploaded_file):
    data = json.loads(uploaded_file.getvalue().decode("utf-8"))
    project_title = data.get("project_title") or data.get("title") or "paper_parallel_project"
    unit = data.get("unit", "段落")

    if "rows" in data:
        rows = [normalize_row(row, i + 1) for i, row in enumerate(data.get("rows", []))]
    elif "alignment_results" in data:
        rows = [normalize_row(row, i + 1, "Imported") for i, row in enumerate(data.get("alignment_results", []))]
    else:
        raise ValueError("rows または alignment_results が見つかりません。")

    glossary = []
    for i, item in enumerate(data.get("glossary", []), start=1):
        glossary.append(
            {
                "id": int(item.get("id", i) or i),
                "english": clean_phrase(item.get("english", item.get("英語", ""))),
                "japanese": clean_phrase(item.get("japanese", item.get("日本語", ""))),
                "category": str(item.get("category", item.get("カテゴリ", "専門用語")) or "専門用語"),
                "note": str(item.get("note", item.get("メモ", ""))),
                "created_at": str(item.get("created_at", "")),
            }
        )

    st.session_state["rows"] = renumber_rows(rows)
    st.session_state["glossary"] = glossary
    cleanup_glossary()
    return project_title, unit, st.session_state["rows"]


def make_json_data(project_title, unit, rows):
    cleanup_glossary()
    data = {
        "app_name": APP_NAME,
        "version": APP_VERSION,
        "saved_at": now_string(),
        "project_title": project_title,
        "unit": unit,
        "rows": rows,
        "glossary": st.session_state.get("glossary", []),
    }
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


# ============================================================
# 表示用データ
# ============================================================
def progress_counts(rows):
    total = len(rows)
    counts = {status: 0 for status in STATUS_OPTIONS}
    for row in rows:
        counts[normalize_status(row.get("状態"))] += 1
    done = counts["確認済み"] + counts["修正済み"]
    percent = round((done / total) * 100, 1) if total else 0.0
    return {
        "total": total,
        "unchecked": counts["未確認"],
        "confirmed": counts["確認済み"],
        "needsfix": counts["要修正"],
        "revised": counts["修正済み"],
        "done": done,
        "percent": percent,
    }


def render_sidebar_progress_card(rows):
    counts = progress_counts(rows)
    total = counts["total"]
    if total <= 0:
        st.markdown(
            """
<div class="ppr-progress-card">
  <div class="ppr-progress-head">
    <span class="ppr-progress-title">進捗</span>
    <span class="ppr-progress-percent">0%</span>
  </div>
</div>
""",
            unsafe_allow_html=True,
        )
        return

    segments = [
        ("ppr-progress-unchecked", counts["unchecked"], "未確認", "未確認の割合"),
        ("ppr-progress-needsfix", counts["needsfix"], "要修正", "要修正の割合"),
        ("ppr-progress-confirmed", counts["confirmed"], "確認済み", "確認済みの割合"),
        ("ppr-progress-revised", counts["revised"], "修正済み", "修正済みの割合"),
    ]

    bar_segments = []
    for class_name, value, label, meaning in segments:
        percent = (value / total * 100) if total else 0
        if value <= 0:
            continue
        bar_segments.append(
            f'''
<div class="ppr-progress-segment {class_name}" style="height:{percent:.4f}%;" title="{html.escape(label, quote=True)}：全体の {percent:.1f}% ／ {value}件">
  <span class="ppr-progress-tip">
    <span class="ppr-progress-tip-label">{html.escape(label)}</span>
    <span class="ppr-progress-tip-detail">{html.escape(meaning)}</span>
    <span class="ppr-progress-tip-detail">全体の {percent:.1f}% ／ {value}件</span>
  </span>
</div>
'''
        )

    st.markdown(
        f'''
<div class="ppr-progress-card">
  <div class="ppr-progress-head">
    <span class="ppr-progress-title">進捗</span>
    <span class="ppr-progress-percent">{counts['percent']}%</span>
  </div>
  <div class="ppr-progress-body">
    <div class="ppr-progress-stack">
      {''.join(bar_segments)}
    </div>
  </div>
</div>
''',
        unsafe_allow_html=True,
    )

def overview_dataframe(rows):
    return pd.DataFrame(
        [
            {
                "ID": row["ID"],
                "章": row["章"],
                "章内ID": row["章内ID"],
                "状態": row["状態"],
                "用語": "あり" if has_glossary_hit(row) else "",
                "英文プレビュー": truncate_text(row.get("英文", ""), 105),
                "訳文プレビュー": truncate_text(get_final_translation(row), 105),
                "メモ": truncate_text(row.get("メモ", ""), 70),
            }
            for row in rows
        ]
    )


def filter_rows(rows, selected_sections, selected_statuses, keyword, only_glossary):
    keyword = str(keyword or "").lower().strip()
    filtered = []
    for row in rows:
        if selected_sections and row.get("章") not in selected_sections:
            continue
        if selected_statuses and row.get("状態") not in selected_statuses:
            continue
        if only_glossary and not has_glossary_hit(row):
            continue
        if keyword:
            haystack = "\n".join(
                [
                    str(row.get("英文", "")),
                    str(row.get("原訳", "")),
                    str(row.get("修正訳", "")),
                    str(row.get("メモ", "")),
                ]
            ).lower()
            if keyword not in haystack:
                continue
        filtered.append(row)
    return filtered


# ============================================================
# ハイライトHTML
# ============================================================
def make_glossary_tooltip(item):
    """用語辞典ハイライト用の小さなHTMLツールチップを作る。

    以前は title / data-tip 属性に改行や <br> 相当の文字列を入れていたため、
    環境によっては「<br>」がそのまま見えてしまうことがあった。
    ここではツールチップ本体をHTML要素として持たせ、行ごとはCSSで縦並びにする。
    """
    rows = []
    english = clean_phrase(item.get("english", ""))
    japanese = clean_phrase(item.get("japanese", ""))
    category = clean_phrase(item.get("category", ""))
    note = str(item.get("note", "")).strip()

    if english:
        rows.append(("英語", english))
    if japanese:
        rows.append(("日本語訳", japanese))
    if category:
        rows.append(("カテゴリ", category))
    if note:
        rows.append(("メモ", note))

    if not rows:
        rows.append(("用語辞典", "登録済み"))

    line_html = []
    for label, value in rows:
        line_html.append(
            '<span class="tooltip-line">'
            f'<span class="tooltip-label">{html.escape(label)}：</span>'
            f'<span class="tooltip-value">{html.escape(value)}</span>'
            '</span>'
        )
    return '<span class="glossary-tooltip" role="tooltip">' + ''.join(line_html) + '</span>'


def highlight_english_terms(text, glossary):
    text = str(text or "")
    terms = []
    for item in glossary:
        term = clean_phrase(item.get("english", ""))
        if not term:
            continue
        terms.append((term, item))

    if not terms:
        return html.escape(text).replace("\n", "<br>")

    terms.sort(key=lambda pair: len(pair[0]), reverse=True)
    pattern = "|".join(f"({make_term_pattern(term)})" for term, _ in terms)

    item_by_lower = {}
    for term, item in terms:
        item_by_lower.setdefault(term.lower(), item)

    result = []
    last = 0
    for match in re.finditer(pattern, text, flags=re.IGNORECASE):
        start, end = match.span()
        if start < last:
            continue
        matched_text = text[start:end]
        result.append(html.escape(text[last:start]))

        key = matched_text.lower()
        item = item_by_lower.get(key)
        if item is None:
            for term, candidate in terms:
                if matched_text.lower() == term.lower():
                    item = candidate
                    break
        tooltip_html = make_glossary_tooltip(item or {})
        result.append(
            f'<span class="hl-glossary">{html.escape(matched_text)}{tooltip_html}</span>'
        )
        last = end

    result.append(html.escape(text[last:]))
    return "".join(result).replace("\n", "<br>")


def build_highlight_html(row, glossary, height=DEFAULT_VIEWER_HEIGHT):
    en_html = highlight_english_terms(row.get("英文", ""), glossary)
    ja_html = html.escape(get_final_translation(row)).replace("\n", "<br>")
    glossary_hits = sum(1 for item in glossary if clean_phrase(item.get("english", "")) and re.search(make_term_pattern(item.get("english", "")), str(row.get("英文", "")), flags=re.IGNORECASE))

    return f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{
    margin:0;
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
    color:#172033;
    background:#ffffff;
}}
.wrap {{ padding: 0; }}
.legend {{
    display:flex;
    flex-wrap:wrap;
    gap:6px;
    margin-bottom:10px;
}}
.badge {{
    display:inline-block;
    padding:5px 10px;
    border-radius:999px;
    border:1px solid #dbe3ef;
    background:#f8fafc;
    font-size:12px;
    font-weight:700;
}}
.viewer {{
    display:grid;
    grid-template-columns:1fr 1fr;
    gap:12px;
    height:{height}px;
}}
.pane {{
    border:1px solid #dbe3ef;
    border-radius:18px;
    padding:16px 18px;
    overflow-y:auto;
    line-height:1.85;
    font-size:15.5px;
    background:linear-gradient(180deg,#ffffff 0%,#fbfdff 100%);
}}
.pane h3 {{
    position:sticky;
    top:-16px;
    z-index:3;
    background:#ffffff;
    padding:12px 0 10px 0;
    margin:-16px 0 12px 0;
    border-bottom:1px solid #eef2f7;
    color:#05285f;
}}
.hl-glossary {{
    position: relative;
    background:#dcfce7;
    border-bottom:2px solid #16a34a;
    border-radius:5px;
    padding:1px 4px;
    cursor:help;
    font-weight: 700;
}}
.glossary-tooltip {{
    display:none;
    position:absolute;
    left:0;
    top:1.9em;
    min-width:240px;
    max-width:360px;
    white-space:normal;
    background:#0f172a;
    color:#ffffff;
    padding:10px 12px;
    border-radius:12px;
    box-shadow:0 10px 24px rgba(15,23,42,.22);
    font-size:12.5px;
    line-height:1.55;
    font-weight:500;
    z-index:1000;
    pointer-events:none;
}}
.hl-glossary:hover .glossary-tooltip {{
    display:block;
}}
.tooltip-line {{
    display:block;
    margin:2px 0;
}}
.tooltip-label {{
    font-weight:800;
    color:#bbf7d0;
}}
.tooltip-value {{
    overflow-wrap:anywhere;
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="legend">
    <span class="badge">緑：用語辞典ハイライト</span>
    <span class="badge">このペアのヒット数：{glossary_hits}</span>
    <span class="badge">カーソルを合わせると訳・カテゴリ・メモを表示</span>
  </div>
  <div class="viewer">
    <div class="pane"><h3>Original English</h3>{en_html}</div>
    <div class="pane"><h3>Japanese Translation</h3>{ja_html}</div>
  </div>
</div>
</body>
</html>
"""


# ============================================================
# 前後文脈・編集操作
# ============================================================
def get_context_rows(rows, row_id, radius):
    index = get_row_index_by_id(rows, row_id)
    if index is None:
        return []
    return rows[max(0, index - radius): min(len(rows), index + radius + 1)]


def render_context(rows, selected_id, radius):
    context = get_context_rows(rows, selected_id, radius)
    for row in context:
        is_current = int(row["ID"]) == int(selected_id)
        bg = "#eef4ff" if is_current else "#ffffff"
        border = "#003f8e" if is_current else "#dbe3ef"
        label = "現在" if is_current else "前後"
        st.markdown(
            f"""
<div style="border:1px solid {border}; background:{bg}; border-radius:14px; padding:12px 14px; margin:8px 0;">
  <div style="font-size:13px; color:#475569; font-weight:800; margin-bottom:6px;">
    {label} ／ ID {row['ID']} ／ {html.escape(str(row.get('章','')))}-{row.get('章内ID')} ／ {html.escape(str(row.get('状態','')))}
  </div>
  <div style="display:grid; grid-template-columns:1fr 1fr; gap:14px;">
    <div style="line-height:1.7;"><b>英文</b><br>{html.escape(str(row.get('英文','')))}</div>
    <div style="line-height:1.7;"><b>訳文</b><br>{html.escape(get_final_translation(row))}</div>
  </div>
</div>
""",
            unsafe_allow_html=True,
        )


def next_sequential_id(rows, current_id):
    index = get_row_index_by_id(rows, current_id)
    if index is None:
        return None
    if index + 1 < len(rows):
        return int(rows[index + 1]["ID"])
    return int(rows[index]["ID"])


def previous_sequential_id(rows, current_id):
    index = get_row_index_by_id(rows, current_id)
    if index is None:
        return None
    if index - 1 >= 0:
        return int(rows[index - 1]["ID"])
    return int(rows[index]["ID"])


def next_unchecked_id(rows, current_id):
    if not rows:
        return None
    index = get_row_index_by_id(rows, current_id)
    if index is None:
        index = -1
    order = list(range(index + 1, len(rows))) + list(range(0, index + 1))
    for i in order:
        if normalize_status(rows[i].get("状態")) in ["未確認", "要修正"]:
            return int(rows[i]["ID"])
    return next_sequential_id(rows, current_id)


def save_row(row_id, edited_translation, status, memo, restore=False):
    rows = st.session_state.get("rows", [])
    index = get_row_index_by_id(rows, row_id)
    if index is None:
        return False
    rows[index]["修正訳"] = str(rows[index].get("原訳", "")) if restore else str(edited_translation)
    rows[index]["状態"] = normalize_status(status)
    rows[index]["メモ"] = str(memo or "")
    rows[index]["更新日時"] = now_string()
    st.session_state["rows"] = rows
    return True


def insert_empty_row_after(rows, row_id):
    index = get_row_index_by_id(rows, row_id)
    if index is None:
        return rows
    base = rows[index]
    rows.insert(
        index + 1,
        {
            "ID": 0,
            "章": base.get("章", "Section"),
            "章内ID": 0,
            "英文": "",
            "原訳": "",
            "修正訳": "",
            "状態": "未確認",
            "メモ": "",
            "更新日時": now_string(),
        },
    )
    return renumber_rows(rows)


def delete_row_by_id(rows, row_id):
    index = get_row_index_by_id(rows, row_id)
    if index is None:
        return rows
    del rows[index]
    return renumber_rows(rows)


# ============================================================
# CSV / Word出力
# ============================================================
def make_csv_data(rows):
    df = pd.DataFrame(
        [
            {
                "ID": row["ID"],
                "章": row["章"],
                "章内ID": row["章内ID"],
                "状態": row["状態"],
                "英文": row.get("英文", ""),
                "原訳": row.get("原訳", ""),
                "修正訳": get_final_translation(row),
                "メモ": row.get("メモ", ""),
                "更新日時": row.get("更新日時", ""),
            }
            for row in rows
        ]
    )
    return df.to_csv(index=False).encode("utf-8-sig")


def make_docx_data(project_title, rows, export_style):
    document = Document()
    document.styles["Normal"].font.name = "Times New Roman"
    document.styles["Normal"].font.size = Pt(10.5)

    document.add_heading(project_title or APP_NAME, level=1)
    document.add_paragraph(f"Exported from {APP_NAME} / {APP_VERSION}")
    document.add_paragraph(f"Saved at: {now_string()}")

    if export_style == "2列表（元文｜訳文）":
        table = document.add_table(rows=1, cols=2)
        table.style = "Table Grid"
        table.rows[0].cells[0].text = "元文"
        table.rows[0].cells[1].text = "訳文"
        for row in rows:
            cells = table.add_row().cells
            cells[0].text = str(row.get("英文", ""))
            cells[1].text = get_final_translation(row)

    elif export_style == "縦並び（元文→訳文）":
        current_section = None
        for row in rows:
            if row.get("章") != current_section:
                current_section = row.get("章")
                document.add_heading(str(current_section), level=2)
            document.add_heading(f"ID {row.get('ID')}", level=3)
            document.add_paragraph("元文").runs[0].bold = True
            document.add_paragraph(str(row.get("英文", "")))
            document.add_paragraph("訳文").runs[0].bold = True
            document.add_paragraph(get_final_translation(row))

    elif export_style == "訳文のみ":
        current_section = None
        for row in rows:
            if row.get("章") != current_section:
                current_section = row.get("章")
                document.add_heading(str(current_section), level=2)
            document.add_paragraph(get_final_translation(row))

    else:
        table = document.add_table(rows=1, cols=6)
        table.style = "Table Grid"
        headers = ["ID", "章", "状態", "元文", "訳文", "メモ"]
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        for row in rows:
            values = [
                row["ID"],
                row["章"],
                row["状態"],
                row.get("英文", ""),
                get_final_translation(row),
                row.get("メモ", ""),
            ]
            cells = table.add_row().cells
            for i, value in enumerate(values):
                cells[i].text = str(value)

        if st.session_state.get("glossary"):
            document.add_page_break()
            document.add_heading("用語辞典", level=2)
            glossary_table = document.add_table(rows=1, cols=4)
            glossary_table.style = "Table Grid"
            for i, header in enumerate(["英語", "日本語", "カテゴリ", "メモ"]):
                glossary_table.rows[0].cells[i].text = header
            for item in st.session_state.get("glossary", []):
                cells = glossary_table.add_row().cells
                cells[0].text = str(item.get("english", ""))
                cells[1].text = str(item.get("japanese", ""))
                cells[2].text = str(item.get("category", ""))
                cells[3].text = str(item.get("note", ""))

    stream = BytesIO()
    document.save(stream)
    stream.seek(0)
    return stream.getvalue()


# ============================================================
# 画面：サイドバー
# ============================================================
def render_sidebar():
    rows = st.session_state.get("rows", [])
    counts = progress_counts(rows)

    with st.sidebar:
        st.header("作業パネル")

        render_sidebar_progress_card(rows)

        st.download_button(
            "JSONを保存",
            data=make_json_data(st.session_state["project_title"], st.session_state["unit"], rows),
            file_name=f"{safe_filename(st.session_state['project_title'])}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json",
            mime="application/json",
            disabled=(not rows and not st.session_state.get("glossary")),
            use_container_width=True,
        )

        st.markdown("---")
        page_options = ["一覧・選択", "読む・修正", "用語辞典", "出力"]
        if st.session_state.get("active_page") not in page_options:
            st.session_state["active_page"] = "一覧・選択"
        st.session_state["active_page"] = st.radio(
            "画面",
            page_options,
            index=page_options.index(st.session_state["active_page"]),
        )

        st.markdown("---")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("設定", use_container_width=True):
                st.session_state["show_settings"] = not st.session_state["show_settings"]
        with c2:
            if st.button("入力", use_container_width=True):
                st.session_state["show_json_loader"] = True
                st.session_state["show_text_adder"] = True

        if st.session_state.get("show_settings", False):
            st.markdown("### 設定")
            st.session_state["project_title"] = st.text_input(
                "プロジェクト名",
                value=st.session_state.get("project_title", "paper_parallel_project"),
            )
            st.session_state["unit"] = st.radio(
                "処理単位",
                ["段落", "文"],
                index=0 if st.session_state.get("unit", "段落") == "段落" else 1,
                help="10000 words級の論文では、まず段落単位がおすすめです。",
            )
            st.session_state["page_size"] = st.selectbox(
                "1ページの表示件数",
                [10, 20, 50, 100, 200],
                index=[10, 20, 50, 100, 200].index(st.session_state.get("page_size", 20)),
            )
            st.session_state["context_radius"] = st.slider(
                "前後文脈の表示範囲",
                min_value=1,
                max_value=5,
                value=int(st.session_state.get("context_radius", 2)),
            )


# ============================================================
# 画面：読み込み・テキスト追加
# ============================================================
def render_input_area():
    show_any = st.session_state.get("show_json_loader", False) or st.session_state.get("show_text_adder", False)
    if not show_any:
        c1, c2 = st.columns(2)
        with c1:
            if st.button("JSON読み込み欄を開く", use_container_width=True):
                st.session_state["show_json_loader"] = True
                st.rerun()
        with c2:
            if st.button("テキスト追加欄を開く", use_container_width=True):
                st.session_state["show_text_adder"] = True
                st.rerun()
        return

    setup_left, setup_right = st.columns(2)

    with setup_left:
        if st.session_state.get("show_json_loader", False):
            with st.expander("保存済みJSONから作業を再開", expanded=True):
                uploaded_json = st.file_uploader(
                    "JSONファイルを選択してください",
                    type=["json"],
                    help="この版のJSONだけでなく、旧 alignment_results 形式もできる範囲で読み込めます。",
                )
                if uploaded_json is not None:
                    st.write(f"選択中：{uploaded_json.name}")
                    if st.button("このJSONを読み込む", type="primary", use_container_width=True):
                        try:
                            project_title, unit, rows = load_rows_from_json(uploaded_json)
                            st.session_state["project_title"] = project_title
                            st.session_state["unit"] = unit if unit in ["段落", "文"] else "段落"
                            st.session_state["selected_id"] = rows[0]["ID"] if rows else None
                            st.session_state["active_page"] = "一覧・選択"
                            st.session_state["show_json_loader"] = False
                            st.session_state["show_text_adder"] = False
                            st.session_state["show_settings"] = False
                            st.success("JSONを読み込みました。入力欄と設定欄は閉じました。")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"読み込みに失敗しました：{exc}")
                if st.button("JSON読み込み欄を閉じる", use_container_width=True):
                    st.session_state["show_json_loader"] = False
                    st.rerun()

    with setup_right:
        if st.session_state.get("show_text_adder", False):
            with st.expander("英文・訳文を追加", expanded=True):
                st.caption(f"現在の分割単位：{st.session_state.get('unit', '段落')}。変更したい場合は左の設定を開いてください。")
                section_title = st.text_input("章・セクション名", value="Section")
                english_text = st.text_area("英文", height=220)
                japanese_text = st.text_area("訳文", height=220)
                add_mode = st.radio(
                    "追加方法",
                    ["現在の作業に追加", "新規作成（現在の作業を置き換え）"],
                    index=0,
                )
                if st.button("テキストを反映", type="primary", use_container_width=True):
                    en_segments = segment_text(english_text, "英語", st.session_state["unit"])
                    ja_segments = segment_text(japanese_text, "日本語", st.session_state["unit"])
                    if not en_segments and not ja_segments:
                        st.warning("英文または訳文を入力してください。")
                    else:
                        existing_rows = [] if add_mode.startswith("新規作成") else st.session_state.get("rows", [])
                        new_rows = make_rows_from_segments(
                            en_segments,
                            ja_segments,
                            section_title,
                            start_id=len(existing_rows) + 1,
                        )
                        st.session_state["rows"] = renumber_rows(existing_rows + new_rows)
                        if st.session_state["rows"] and st.session_state.get("selected_id") is None:
                            st.session_state["selected_id"] = st.session_state["rows"][0]["ID"]
                        st.session_state["active_page"] = "一覧・選択"
                        st.session_state["show_json_loader"] = False
                        st.session_state["show_text_adder"] = False
                        st.session_state["show_settings"] = False
                        st.success("テキストを追加しました。入力欄と設定欄は閉じました。")
                        st.rerun()
                if st.button("テキスト追加欄を閉じる", use_container_width=True):
                    st.session_state["show_text_adder"] = False
                    st.rerun()


# ============================================================
# 画面：一覧・選択
# ============================================================
def render_overview_page():
    rows = st.session_state.get("rows", [])
    if not rows:
        st.info("まずJSONを読み込むか、英文・訳文を追加してください。")
        return

    st.markdown("### 一覧・選択")
    filter_left, filter_mid, filter_right = st.columns([1, 1, 1])
    sections = sorted({str(row.get("章", "")) for row in rows})
    with filter_left:
        selected_sections = st.multiselect("章で絞り込み", sections)
    with filter_mid:
        selected_statuses = st.multiselect("状態で絞り込み", STATUS_OPTIONS)
    with filter_right:
        keyword = st.text_input("キーワード検索")
    only_glossary = st.checkbox("用語辞典に登録した英語が出てくるペアだけ表示")

    filtered = filter_rows(rows, selected_sections, selected_statuses, keyword, only_glossary)
    st.session_state["last_filtered_ids"] = [int(row["ID"]) for row in filtered]

    st.markdown(f"<span class='ppr-chip ppr-chip-blue'>表示対象：{len(filtered)} / {len(rows)}</span>", unsafe_allow_html=True)

    if not filtered:
        st.warning("条件に合うペアがありません。")
        return

    page_size = int(st.session_state.get("page_size", 20))
    total_pages = max(1, math.ceil(len(filtered) / page_size))
    page_number = st.number_input("ページ", min_value=1, max_value=total_pages, value=1, step=1)
    start = (int(page_number) - 1) * page_size
    end = start + page_size
    page_rows = filtered[start:end]

    overview_df = overview_dataframe(page_rows)
    table_event = st.dataframe(
        overview_df,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
        key=f"overview_table_{int(page_number)}_{page_size}_{len(filtered)}",
    )
    selected_table_rows = []
    if hasattr(table_event, "selection") and hasattr(table_event.selection, "rows"):
        selected_table_rows = table_event.selection.rows
    elif isinstance(table_event, dict):
        selected_table_rows = table_event.get("selection", {}).get("rows", [])

    if selected_table_rows:
        picked_index = int(selected_table_rows[0])
        if 0 <= picked_index < len(overview_df):
            st.session_state["selected_id"] = int(overview_df.iloc[picked_index]["ID"])

    choices = [int(row["ID"]) for row in page_rows]
    current_id = st.session_state.get("selected_id")
    default_index = choices.index(current_id) if current_id in choices else 0
    selected_id = st.selectbox(
        "読むペアを選ぶ",
        choices,
        index=default_index,
        format_func=lambda rid: f"ID {rid}: {truncate_text(get_row_by_id(rows, rid).get('英文', ''), 80)}",
    )
    if st.button("選択して読む", type="primary"):
        st.session_state["selected_id"] = selected_id
        st.session_state["active_page"] = "読む・修正"
        st.rerun()


# ============================================================
# 画面：読む・修正
# ============================================================
def render_read_edit_page():
    rows = st.session_state.get("rows", [])
    if not rows:
        st.info("まずJSONを読み込むか、英文・訳文を追加してください。")
        return

    if st.session_state.get("selected_id") is None:
        st.session_state["selected_id"] = rows[0]["ID"]

    selected_row = get_row_by_id(rows, st.session_state["selected_id"])
    if selected_row is None:
        st.session_state["selected_id"] = rows[0]["ID"]
        selected_row = rows[0]

    st.markdown("### 読む・修正")
    current_index = get_row_index_by_id(rows, selected_row["ID"])

    top_left, top_right = st.columns([1.2, 1])
    with top_left:
        glossary_hit = has_glossary_hit(selected_row)
        st.markdown(
            f"""
<div class="ppr-card">
  <b>選択中：</b>ID {selected_row['ID']} ／ {html.escape(str(selected_row.get('章','')))}-{selected_row.get('章内ID')}<br>
  <span class="ppr-chip ppr-chip-blue">状態：{html.escape(str(selected_row.get('状態','')))}</span>
  <span class="ppr-chip ppr-chip-green">用語辞典ヒット：{'あり' if glossary_hit else 'なし'}</span>
</div>
""",
            unsafe_allow_html=True,
        )
    with top_right:
        n1, n2, n3, n4 = st.columns(4)
        with n1:
            if st.button("前へ", disabled=(current_index is None or current_index <= 0), use_container_width=True):
                st.session_state["selected_id"] = previous_sequential_id(rows, selected_row["ID"])
                st.rerun()
        with n2:
            if st.button("次へ", disabled=(current_index is None or current_index >= len(rows) - 1), use_container_width=True):
                st.session_state["selected_id"] = next_sequential_id(rows, selected_row["ID"])
                st.rerun()
        with n3:
            if st.button("次の未確認", use_container_width=True):
                st.session_state["selected_id"] = next_unchecked_id(rows, selected_row["ID"])
                st.rerun()
        with n4:
            if st.button("一覧へ", use_container_width=True):
                st.session_state["active_page"] = "一覧・選択"
                st.rerun()

    left, right = st.columns([1.2, 1])
    with left:
        components.html(
            build_highlight_html(selected_row, st.session_state.get("glossary", [])),
            height=DEFAULT_VIEWER_HEIGHT + 76,
            scrolling=False,
        )
        with st.expander("前後文脈を見る", expanded=False):
            render_context(rows, selected_row["ID"], int(st.session_state.get("context_radius", 2)))

    with right:
        with st.form(key=f"edit_form_{selected_row['ID']}"):
            st.text_area("英文（確認用）", value=str(selected_row.get("英文", "")), height=160, disabled=True)
            edited_translation = st.text_area("修正訳", value=get_final_translation(selected_row), height=230)
            status = st.selectbox(
                "状態",
                STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(normalize_status(selected_row.get("状態"))),
            )
            memo = st.text_area("メモ", value=str(selected_row.get("メモ", "")), height=80)
            restore = st.checkbox("原訳に戻す")

            save_only = st.form_submit_button("このペアを保存")
            confirm_and_next = st.form_submit_button("確認済みにして次へ", type="primary")
            revise_and_next = st.form_submit_button("修正済みにして次へ")

        if save_only or confirm_and_next or revise_and_next:
            save_status = status
            if confirm_and_next:
                save_status = "確認済み"
            elif revise_and_next:
                save_status = "修正済み"

            ok = save_row(selected_row["ID"], edited_translation, save_status, memo, restore)
            if ok:
                if confirm_and_next or revise_and_next:
                    st.session_state["selected_id"] = next_sequential_id(st.session_state["rows"], selected_row["ID"])
                st.success("保存しました。")
                st.rerun()
            else:
                st.error("保存できませんでした。")

        with st.expander("結合・削除・空ペア", expanded=False):
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("空ペア追加", use_container_width=True):
                    st.session_state["rows"] = insert_empty_row_after(rows, selected_row["ID"])
                    st.session_state["selected_id"] = next_sequential_id(st.session_state["rows"], selected_row["ID"])
                    st.rerun()
            with c2:
                if st.button("次の訳文を結合", disabled=(current_index is None or current_index >= len(rows) - 1), use_container_width=True):
                    current = rows[current_index]
                    next_row = rows[current_index + 1]
                    current["原訳"] = "\n\n".join(
                        [p for p in [str(current.get("原訳", "")).strip(), str(next_row.get("原訳", "")).strip()] if p]
                    )
                    current["修正訳"] = "\n\n".join(
                        [p for p in [get_final_translation(current), get_final_translation(next_row)] if p]
                    )
                    current["状態"] = "修正済み"
                    current["更新日時"] = now_string()
                    del rows[current_index + 1]
                    st.session_state["rows"] = renumber_rows(rows)
                    st.success("次の訳文を結合しました。")
                    st.rerun()
            with c3:
                if st.button("このペアを削除", disabled=(len(rows) <= 1), use_container_width=True):
                    next_id = next_sequential_id(rows, selected_row["ID"])
                    st.session_state["rows"] = delete_row_by_id(rows, selected_row["ID"])
                    if st.session_state["rows"]:
                        st.session_state["selected_id"] = next_id if get_row_by_id(st.session_state["rows"], next_id) else st.session_state["rows"][0]["ID"]
                    else:
                        st.session_state["selected_id"] = None
                    st.rerun()


# ============================================================
# 画面：用語辞典
# ============================================================
def render_glossary_page():
    st.markdown("### 用語辞典")
    st.caption("英語用語を登録すると、読む・修正画面の英文中で自動ハイライトされます。カーソルを合わせると訳・カテゴリ・メモが出ます。")

    with st.form("glossary_add_form", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            english = st.text_input("英語")
        with c2:
            japanese = st.text_input("日本語訳")
        c3, c4 = st.columns([1, 2])
        with c3:
            category = st.selectbox("カテゴリ", GLOSSARY_CATEGORIES)
        with c4:
            note = st.text_input("メモ")
        submitted = st.form_submit_button("用語を追加", type="primary")

    if submitted:
        ok, message = add_unique_glossary(english, japanese, category, note)
        if ok:
            st.success(message)
            st.rerun()
        else:
            st.warning(message)

    glossary = st.session_state.get("glossary", [])
    if not glossary:
        st.info("まだ用語が登録されていません。")
        return

    st.markdown("#### 登録済み用語")
    df = pd.DataFrame(glossary)
    display_columns = ["english", "japanese", "category", "note"]
    edited_df = st.data_editor(
        df[display_columns],
        use_container_width=True,
        num_rows="dynamic",
        hide_index=True,
        column_config={
            "english": "英語",
            "japanese": "日本語訳",
            "category": st.column_config.SelectboxColumn("カテゴリ", options=GLOSSARY_CATEGORIES),
            "note": "メモ",
        },
    )
    if st.button("用語辞典の編集を保存", type="primary"):
        new_glossary = []
        for _, row in edited_df.iterrows():
            english_text = clean_phrase(row.get("english", ""))
            japanese_text = clean_phrase(row.get("japanese", ""))
            if not english_text and not japanese_text:
                continue
            new_glossary.append(
                {
                    "id": len(new_glossary) + 1,
                    "english": english_text,
                    "japanese": japanese_text,
                    "category": str(row.get("category", "専門用語") or "専門用語"),
                    "note": str(row.get("note", "")),
                    "created_at": now_string(),
                }
            )
        st.session_state["glossary"] = new_glossary
        cleanup_glossary()
        st.success("用語辞典を保存しました。")
        st.rerun()


# ============================================================
# 画面：出力
# ============================================================
def render_export_page():
    rows = st.session_state.get("rows", [])
    st.markdown("### 出力")
    if not rows:
        st.info("出力するデータがまだありません。")
        return

    base_name = safe_filename(st.session_state.get("project_title", "paper_parallel"))
    c1, c2 = st.columns(2)
    with c1:
        st.download_button(
            "CSVを保存",
            data=make_csv_data(rows),
            file_name=f"{base_name}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with c2:
        export_style = st.selectbox("Word出力形式", WORD_EXPORT_OPTIONS)
        st.download_button(
            "Wordを保存",
            data=make_docx_data(st.session_state.get("project_title"), rows, export_style),
            file_name=f"{base_name}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )

    st.caption("JSON保存は左の作業パネルからいつでもできます。")


# ============================================================
# アプリ本体
# ============================================================
inject_style()
init_state()
cleanup_glossary()
st.markdown(
    f"""
<div class="ppr-hero">
  <h1>Paper Parallel Reader</h1>
  <p>用語辞典を中心に、論文の英文と和訳を軽く確認・修正するローカルエディタ</p>
  <span class="ppr-badge">Version: {APP_VERSION}</span>
</div>
""",
    unsafe_allow_html=True,
)

render_sidebar()
render_input_area()

active_page = st.session_state.get("active_page", "一覧・選択")
if active_page == "一覧・選択":
    render_overview_page()
elif active_page == "読む・修正":
    render_read_edit_page()
elif active_page == "用語辞典":
    render_glossary_page()
elif active_page == "出力":
    render_export_page()
