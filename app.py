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
APP_VERSION = "1.4.2-page-nav-fix"

STATUS_OPTIONS = ["未確認", "確認済み", "要修正", "修正済み"]
MARKER_LABELS = ["要確認", "不自然かも", "用語確認", "原文確認", "修正済み", "重要"]
LINK_LABELS = ["意味対応", "専門用語", "句・節対応", "要確認", "主張の核"]
GLOSSARY_CATEGORIES = ["専門用語", "理論概念", "方法", "固有名詞", "その他"]
WORD_EXPORT_OPTIONS = [
    "2列表（元文｜訳文）",
    "縦並び（元文→訳文）",
    "訳文のみ",
    "確認用（マーカー・リンクつき）",
]

st.set_page_config(page_title=APP_NAME, layout="wide")


# ============================================================
# デザイン
# ============================================================
def inject_style():
    st.markdown(
        """
<style>
:root {
  --main-blue: #003f8e;
  --deep-blue: #05285f;
  --soft-blue: #eaf2ff;
  --main-green: #16a34a;
  --soft-green: #e8f8ee;
  --paper: #f7f9fc;
  --ink: #172033;
  --muted: #64748b;
  --line: #dbe3ef;
}

.block-container {
  padding-top: 1.3rem;
  padding-bottom: 3rem;
}

.ppr-hero {
  background: linear-gradient(135deg, var(--deep-blue) 0%, var(--main-blue) 62%, #0d9488 100%);
  color: white;
  border-radius: 22px;
  padding: 22px 26px;
  box-shadow: 0 14px 34px rgba(0, 63, 142, .24);
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
.ppr-chip-blue { background: var(--soft-blue); border-color:#bfdbfe; color:#1d4ed8; }
.ppr-chip-green { background: var(--soft-green); border-color:#bbf7d0; color:#15803d; }
.ppr-chip-yellow { background:#fff7d6; border-color:#fde68a; color:#92400e; }

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
  background: linear-gradient(180deg, #f8fbff 0%, #f2fbf5 100%);
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
        "markers": [],
        "links": [],
        "project_title": "paper_parallel_project",
        "unit": "段落",
        "selected_id": None,
        "active_page": "一覧・選択",
        "_active_page_widget": "一覧・選択",
        "_pending_active_page": None,
        "last_filtered_ids": [],
        "export_cache": {},
        "_marker_index": {},
        "_link_index": {},
        "_marked_ids": set(),
        "_linked_ids": set(),
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


def next_item_id(items):
    ids = []
    for item in items:
        try:
            ids.append(int(item.get("id", 0)))
        except Exception:
            pass
    return max(ids or [0]) + 1


def clean_phrase(text):
    return re.sub(r"\s+", " ", str(text or "")).strip()


def lower_key(text):
    return clean_phrase(text).lower()


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
        original = str(row.get("原訳", row.get("和訳", "")))
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
    old_to_new = {}
    for index, row in enumerate(rows, start=1):
        old_id = int(row.get("ID", index) or index)
        section = str(row.get("章", "Section") or "Section")
        section_counts[section] = section_counts.get(section, 0) + 1
        row["ID"] = index
        row["章内ID"] = section_counts[section]
        old_to_new[old_id] = index

    for collection_name in ["markers", "links"]:
        if collection_name in st.session_state:
            fixed = []
            for item in st.session_state[collection_name]:
                try:
                    old_row_id = int(item.get("row_id", -1))
                    if old_row_id in old_to_new:
                        item["row_id"] = old_to_new[old_row_id]
                        fixed.append(item)
                except Exception:
                    pass
            st.session_state[collection_name] = fixed
    return rows


# ============================================================
# 重複対策
# ============================================================
def marker_unique_key(marker):
    return (
        int(marker.get("row_id", -1)),
        clean_phrase(marker.get("ja_phrase", "")),
        str(marker.get("label", "")),
    )


def link_unique_key(link):
    return (
        int(link.get("row_id", -1)),
        lower_key(link.get("ja_phrase", "")),
        lower_key(link.get("en_phrase", "")),
        str(link.get("label", "")),
    )


def glossary_unique_key(item):
    return (
        lower_key(item.get("english", "")),
        clean_phrase(item.get("japanese", "")),
        str(item.get("category", "")),
    )


def dedupe_list(items, key_func):
    seen = set()
    result = []
    for item in items:
        key = key_func(item)
        if key in seen:
            continue
        seen.add(key)
        result.append(item)
    for idx, item in enumerate(result, start=1):
        item["id"] = idx
    return result


def cleanup_state():
    valid_row_ids = {int(row["ID"]) for row in st.session_state.get("rows", [])}

    markers = []
    for item in st.session_state.get("markers", []):
        try:
            row_id = int(item.get("row_id", -1))
        except Exception:
            continue
        phrase = clean_phrase(item.get("ja_phrase", ""))
        if row_id in valid_row_ids and phrase:
            markers.append(
                {
                    "id": int(item.get("id", len(markers) + 1) or len(markers) + 1),
                    "row_id": row_id,
                    "ja_phrase": phrase,
                    "label": str(item.get("label", "要確認") or "要確認"),
                    "note": str(item.get("note", "")),
                    "created_at": str(item.get("created_at", "")),
                }
            )
    st.session_state["markers"] = dedupe_list(markers, marker_unique_key)

    links = []
    for item in st.session_state.get("links", []):
        try:
            row_id = int(item.get("row_id", -1))
        except Exception:
            continue
        ja_phrase = clean_phrase(item.get("ja_phrase", ""))
        en_phrase = clean_phrase(item.get("en_phrase", ""))
        if row_id in valid_row_ids and ja_phrase and en_phrase:
            links.append(
                {
                    "id": int(item.get("id", len(links) + 1) or len(links) + 1),
                    "row_id": row_id,
                    "ja_phrase": ja_phrase,
                    "en_phrase": en_phrase,
                    "label": str(item.get("label", "意味対応") or "意味対応"),
                    "note": str(item.get("note", "")),
                    "created_at": str(item.get("created_at", "")),
                }
            )
    st.session_state["links"] = dedupe_list(links, link_unique_key)

    glossary = []
    for item in st.session_state.get("glossary", []):
        english = clean_phrase(item.get("english", item.get("英語", "")))
        japanese = clean_phrase(item.get("japanese", item.get("日本語", "")))
        if english or japanese:
            glossary.append(
                {
                    "id": int(item.get("id", len(glossary) + 1) or len(glossary) + 1),
                    "english": english,
                    "japanese": japanese,
                    "category": str(item.get("category", item.get("カテゴリ", "専門用語")) or "専門用語"),
                    "note": str(item.get("note", item.get("メモ", ""))),
                    "created_at": str(item.get("created_at", "")),
                }
            )
    st.session_state["glossary"] = dedupe_list(glossary, glossary_unique_key)
    rebuild_lookup_indexes()


def add_unique_marker(row_id, phrase, label, note):
    phrase = clean_phrase(phrase)
    if not phrase:
        return False, "日本語句を入力してください。"
    marker = {
        "id": next_item_id(st.session_state["markers"]),
        "row_id": int(row_id),
        "ja_phrase": phrase,
        "label": label,
        "note": str(note or "").strip(),
        "created_at": now_string(),
    }
    key = marker_unique_key(marker)
    existing_keys = {marker_unique_key(m) for m in st.session_state["markers"]}
    if key in existing_keys:
        return False, "同じマーカーはすでに登録されています。"
    st.session_state["markers"].append(marker)
    cleanup_state()
    return True, "日本語マーカーを追加しました。"


def add_unique_link(row_id, ja_phrase, en_phrase, label, note):
    ja_phrase = clean_phrase(ja_phrase)
    en_phrase = clean_phrase(en_phrase)
    if not ja_phrase or not en_phrase:
        return False, "日本語句と英語句の両方を入力してください。"
    link = {
        "id": next_item_id(st.session_state["links"]),
        "row_id": int(row_id),
        "ja_phrase": ja_phrase,
        "en_phrase": en_phrase,
        "label": label,
        "note": str(note or "").strip(),
        "created_at": now_string(),
    }
    key = link_unique_key(link)
    existing_keys = {link_unique_key(l) for l in st.session_state["links"]}
    if key in existing_keys:
        return False, "同じ英日リンクはすでに登録されています。"
    st.session_state["links"].append(link)
    cleanup_state()
    return True, "英日リンクを追加しました。"


def add_unique_glossary(english, japanese, category, note):
    english = clean_phrase(english)
    japanese = clean_phrase(japanese)
    if not english and not japanese:
        return False, "英語用語または日本語訳を入力してください。"
    item = {
        "id": next_item_id(st.session_state["glossary"]),
        "english": english,
        "japanese": japanese,
        "category": category,
        "note": str(note or "").strip(),
        "created_at": now_string(),
    }
    key = glossary_unique_key(item)
    existing_keys = {glossary_unique_key(g) for g in st.session_state["glossary"]}
    if key in existing_keys:
        return False, "同じ用語はすでに登録されています。"
    st.session_state["glossary"].append(item)
    cleanup_state()
    return True, "用語辞書に追加しました。"


# ============================================================
# JSON読み込み
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

    rows = renumber_rows(rows)
    valid_ids = {int(row["ID"]) for row in rows}

    markers = []
    for i, item in enumerate(data.get("markers", []), start=1):
        try:
            row_id = int(item.get("row_id", 1))
        except Exception:
            row_id = 1
        if row_id in valid_ids:
            markers.append(
                {
                    "id": int(item.get("id", i) or i),
                    "row_id": row_id,
                    "ja_phrase": clean_phrase(item.get("ja_phrase", item.get("日本語句", ""))),
                    "label": str(item.get("label", item.get("ラベル", "要確認")) or "要確認"),
                    "note": str(item.get("note", item.get("メモ", ""))),
                    "created_at": str(item.get("created_at", "")),
                }
            )

    links = []
    for i, item in enumerate(data.get("links", []), start=1):
        try:
            row_id = int(item.get("row_id", 1))
        except Exception:
            row_id = 1
        if row_id in valid_ids:
            links.append(
                {
                    "id": int(item.get("id", i) or i),
                    "row_id": row_id,
                    "ja_phrase": clean_phrase(item.get("ja_phrase", item.get("日本語句", ""))),
                    "en_phrase": clean_phrase(item.get("en_phrase", item.get("英語句", ""))),
                    "label": str(item.get("label", item.get("ラベル", "意味対応")) or "意味対応"),
                    "note": str(item.get("note", item.get("メモ", ""))),
                    "created_at": str(item.get("created_at", "")),
                }
            )

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

    st.session_state["rows"] = rows
    st.session_state["markers"] = markers
    st.session_state["links"] = links
    st.session_state["glossary"] = glossary
    cleanup_state()
    return project_title, unit, st.session_state["rows"]


# ============================================================
# 抽出・候補
# ============================================================
def rebuild_lookup_indexes():
    """マーカー・リンクを row_id ごとに引けるようにする。

    行ごとに毎回リスト全体を走査すると、マーカーやリンクが増えた時に重くなる。
    ここで辞書化しておくことで、markers_for(row_id) / links_for(row_id) を高速化する。
    """
    marker_index = {}
    for marker in st.session_state.get("markers", []):
        try:
            row_id = int(marker.get("row_id", -1))
        except Exception:
            continue
        marker_index.setdefault(row_id, []).append(marker)

    link_index = {}
    for link in st.session_state.get("links", []):
        try:
            row_id = int(link.get("row_id", -1))
        except Exception:
            continue
        link_index.setdefault(row_id, []).append(link)

    st.session_state["_marker_index"] = marker_index
    st.session_state["_link_index"] = link_index
    st.session_state["_marked_ids"] = set(marker_index.keys())
    st.session_state["_linked_ids"] = set(link_index.keys())


def markers_for(row_id):
    try:
        row_id = int(row_id)
    except Exception:
        return []
    if "_marker_index" not in st.session_state:
        rebuild_lookup_indexes()
    return st.session_state.get("_marker_index", {}).get(row_id, [])


def links_for(row_id):
    try:
        row_id = int(row_id)
    except Exception:
        return []
    if "_link_index" not in st.session_state:
        rebuild_lookup_indexes()
    return st.session_state.get("_link_index", {}).get(row_id, [])


def marked_row_ids():
    if "_marked_ids" not in st.session_state:
        rebuild_lookup_indexes()
    return st.session_state.get("_marked_ids", set())


def linked_row_ids():
    if "_linked_ids" not in st.session_state:
        rebuild_lookup_indexes()
    return st.session_state.get("_linked_ids", set())


def has_glossary_hit(row):
    en = str(row.get("英文", ""))
    ja = get_final_translation(row)
    for item in st.session_state["glossary"]:
        english = str(item.get("english", "")).strip()
        japanese = str(item.get("japanese", "")).strip()
        if english and english.lower() in en.lower():
            return True
        if japanese and japanese in ja:
            return True
    return False


def suggest_japanese_phrases(text, limit=16):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    pieces = re.split(r"[。！？!?、，,；;：:]", text)
    result = []
    seen = set()
    for piece in pieces:
        piece = piece.strip()
        if 3 <= len(piece) <= 55 and piece not in seen:
            seen.add(piece)
            result.append(piece)
    return result[:limit]


def suggest_english_phrases(text, limit=16):
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    result = []
    seen = set()

    for piece in re.split(r"[,;:]", text):
        piece = piece.strip(" .()[]")
        if 2 <= len(piece.split()) <= 14 and piece.lower() not in seen:
            seen.add(piece.lower())
            result.append(piece)

    words = text.split()
    for size in [3, 4, 5, 6]:
        step = max(1, size - 2)
        for i in range(0, max(0, len(words) - size + 1), step):
            phrase = " ".join(words[i:i + size]).strip(" ,.;:()[]")
            if 2 <= len(phrase.split()) <= 8 and phrase.lower() not in seen:
                seen.add(phrase.lower())
                result.append(phrase)
            if len(result) >= limit:
                return result[:limit]
    return result[:limit]


def filter_rows(rows, selected_sections, selected_statuses, keyword, only_marked, only_linked, only_glossary):
    keyword = str(keyword or "").lower().strip()
    marked_ids = marked_row_ids()
    linked_ids = linked_row_ids()

    filtered = []
    for row in rows:
        row_id = int(row["ID"])
        if selected_sections and row.get("章") not in selected_sections:
            continue
        if selected_statuses and row.get("状態") not in selected_statuses:
            continue
        if only_marked and row_id not in marked_ids:
            continue
        if only_linked and row_id not in linked_ids:
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


def overview_dataframe(rows):
    marked_ids = marked_row_ids()
    linked_ids = linked_row_ids()
    return pd.DataFrame(
        [
            {
                "ID": row["ID"],
                "章": row["章"],
                "章内ID": row["章内ID"],
                "状態": row["状態"],
                "マーカー": "●" if int(row["ID"]) in marked_ids else "",
                "英日リンク": "●" if int(row["ID"]) in linked_ids else "",
                "英文プレビュー": truncate_text(row.get("英文", ""), 105),
                "訳文プレビュー": truncate_text(get_final_translation(row), 105),
                "メモ": truncate_text(row.get("メモ", ""), 70),
            }
            for row in rows
        ]
    )


# ============================================================
# ハイライトHTML
# ============================================================
def build_highlight_specs(row):
    row_id = int(row["ID"])
    en_specs = []
    ja_specs = []

    for link in links_for(row_id):
        link_key = f"link-{link['id']}"
        en_specs.append(
            {
                "phrase": link["en_phrase"],
                "class": "hl-link",
                "data": f'data-link="{link_key}"',
                "title": f"英日リンク：{link.get('label', '')}",
                "priority": 4,
            }
        )
        ja_specs.append(
            {
                "phrase": link["ja_phrase"],
                "class": "hl-link",
                "data": f'data-link="{link_key}"',
                "title": f"英日リンク：{link.get('label', '')}",
                "priority": 4,
            }
        )

    for marker in markers_for(row_id):
        ja_specs.append(
            {
                "phrase": marker["ja_phrase"],
                "class": "hl-marker",
                "data": "",
                "title": f"日本語マーカー：{marker.get('label', '')}",
                "priority": 3,
            }
        )

    for item in st.session_state["glossary"]:
        if item.get("english"):
            en_specs.append(
                {
                    "phrase": item["english"],
                    "class": "hl-glossary",
                    "data": "",
                    "title": f"用語辞書：{item.get('category', '')}",
                    "priority": 2,
                }
            )
        if item.get("japanese"):
            ja_specs.append(
                {
                    "phrase": item["japanese"],
                    "class": "hl-glossary",
                    "data": "",
                    "title": f"用語辞書：{item.get('category', '')}",
                    "priority": 2,
                }
            )
    return en_specs, ja_specs


def apply_highlights(text, specs):
    result = html.escape(str(text or ""))
    specs = [s for s in specs if clean_phrase(s.get("phrase", ""))]
    specs.sort(key=lambda s: (int(s.get("priority", 0)), len(str(s.get("phrase", "")))), reverse=True)

    placeholders = {}
    for i, spec in enumerate(specs):
        phrase = clean_phrase(spec["phrase"])
        escaped_phrase = html.escape(phrase)
        if not escaped_phrase or escaped_phrase not in result:
            continue
        placeholder = f"@@PPR_PLACEHOLDER_{i}_{len(placeholders)}@@"
        span = (
            f'<span class="{spec["class"]}" {spec.get("data", "")} '
            f'title="{html.escape(spec.get("title", ""))}">{escaped_phrase}</span>'
        )
        result = result.replace(escaped_phrase, placeholder)
        placeholders[placeholder] = span

    for placeholder, span in placeholders.items():
        result = result.replace(placeholder, span)
    return result.replace("\n", "<br>")


def build_highlight_html(row, height):
    en_specs, ja_specs = build_highlight_specs(row)
    en_html = apply_highlights(row.get("英文", ""), en_specs)
    ja_html = apply_highlights(get_final_translation(row), ja_specs)
    marker_count = len(markers_for(row["ID"]))
    link_count = len(links_for(row["ID"]))

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
.hl-marker {{
    background:#fff2a8;
    border-bottom:2px solid #d7a900;
    border-radius:5px;
    padding:1px 4px;
}}
.hl-link {{
    background:#dbeafe;
    border-bottom:2px solid #2563eb;
    border-radius:5px;
    padding:1px 4px;
    cursor:pointer;
}}
.hl-glossary {{
    background:#dcfce7;
    border-bottom:2px solid #16a34a;
    border-radius:5px;
    padding:1px 4px;
}}
.hl-link.active {{
    background:#bfdbfe;
    outline:3px solid #1d4ed8;
    box-shadow:0 0 0 4px rgba(37,99,235,.16);
}}
</style>
</head>
<body>
<div class="wrap">
  <div class="legend">
    <span class="badge">黄：日本語マーカー {marker_count}件</span>
    <span class="badge">青：英日リンク {link_count}件</span>
    <span class="badge">緑：用語辞書</span>
    <span class="badge">青い語句をクリックで対応箇所を強調</span>
  </div>
  <div class="viewer">
    <div class="pane"><h3>Original English</h3>{en_html}</div>
    <div class="pane"><h3>Japanese Translation</h3>{ja_html}</div>
  </div>
</div>
<script>
let activeLink = null;
function clearActive() {{
  document.querySelectorAll('.hl-link').forEach(el => el.classList.remove('active'));
}}
document.querySelectorAll('.hl-link').forEach(el => {{
  el.addEventListener('click', function(event) {{
    event.stopPropagation();
    const linkId = this.dataset.link;
    if (!linkId) return;
    if (activeLink === linkId) {{
      activeLink = null;
      clearActive();
      return;
    }}
    activeLink = linkId;
    clearActive();
    document.querySelectorAll(`[data-link="${{linkId}}"]`).forEach(target => {{
      target.classList.add('active');
      target.scrollIntoView({{behavior:'smooth', block:'center'}});
    }});
  }});
}});
document.body.addEventListener('click', () => {{
  activeLink = null;
  clearActive();
}});
</script>
</body>
</html>
"""


# ============================================================
# 前後文脈・保存
# ============================================================
def get_context_rows(rows, row_id, radius):
    index = get_row_index_by_id(rows, row_id)
    if index is None:
        return []
    return rows[max(0, index - radius): min(len(rows), index + radius + 1)]


def render_context(rows, selected_id, radius):
    context = get_context_rows(rows, selected_id, radius)
    if not context:
        return
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


def make_json_data(project_title, unit, rows):
    cleanup_state()
    data = {
        "app_name": APP_NAME,
        "version": APP_VERSION,
        "saved_at": now_string(),
        "project_title": project_title,
        "unit": unit,
        "rows": rows,
        "glossary": st.session_state["glossary"],
        "markers": st.session_state["markers"],
        "links": st.session_state["links"],
    }
    return json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")


def make_csv_data(rows):
    marked_ids = marked_row_ids()
    linked_ids = linked_row_ids()
    df = pd.DataFrame(
        [
            {
                "ID": row["ID"],
                "章": row["章"],
                "章内ID": row["章内ID"],
                "状態": row["状態"],
                "マーカーあり": "yes" if int(row["ID"]) in marked_ids else "",
                "英日リンクあり": "yes" if int(row["ID"]) in linked_ids else "",
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


def marker_summary(row_id):
    return " / ".join([f"{m['label']}: {m['ja_phrase']}" for m in markers_for(row_id)])


def link_summary(row_id):
    return " / ".join([f"{l['ja_phrase']} ⇔ {l['en_phrase']}" for l in links_for(row_id)])


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
            document.add_heading(f"ID {row['ID']}", level=3)
            p1 = document.add_paragraph()
            p1.add_run("元文").bold = True
            document.add_paragraph(str(row.get("英文", "")))
            p2 = document.add_paragraph()
            p2.add_run("訳文").bold = True
            document.add_paragraph(get_final_translation(row))

    elif export_style == "訳文のみ":
        current_section = None
        for row in rows:
            if row.get("章") != current_section:
                current_section = row.get("章")
                document.add_heading(str(current_section), level=2)
            document.add_paragraph(get_final_translation(row))

    else:
        table = document.add_table(rows=1, cols=8)
        table.style = "Table Grid"
        headers = ["ID", "章", "状態", "元文", "訳文", "メモ", "マーカー", "英日リンク"]
        for i, header in enumerate(headers):
            table.rows[0].cells[i].text = header
        for row in rows:
            row_id = int(row["ID"])
            values = [
                row["ID"],
                row["章"],
                row["状態"],
                row.get("英文", ""),
                get_final_translation(row),
                row.get("メモ", ""),
                marker_summary(row_id),
                link_summary(row_id),
            ]
            cells = table.add_row().cells
            for i, value in enumerate(values):
                cells[i].text = str(value)

        if st.session_state["glossary"]:
            document.add_page_break()
            document.add_heading("用語辞書", level=2)
            glossary_table = document.add_table(rows=1, cols=4)
            glossary_table.style = "Table Grid"
            for i, header in enumerate(["英語", "日本語", "カテゴリ", "メモ"]):
                glossary_table.rows[0].cells[i].text = header
            for item in st.session_state["glossary"]:
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
# アプリ本体
# ============================================================
inject_style()
init_state()
cleanup_state()

st.markdown(
    f"""
<div class="ppr-hero">
  <h1>📚 Paper Parallel Reader</h1>
  <p>前後文脈を見ながら、和訳を自然に整えるための軽量ローカルエディタ</p>
  <span class="ppr-badge">Version: {APP_VERSION}</span>
  <span class="ppr-badge">Marker</span>
  <span class="ppr-badge">EN-JA Link</span>
  <span class="ppr-badge">Glossary</span>
</div>
""",
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("設定")
    st.session_state["project_title"] = st.text_input("プロジェクト名", value=st.session_state["project_title"])
    st.session_state["unit"] = st.radio(
        "処理単位",
        ["段落", "文"],
        index=0 if st.session_state["unit"] == "段落" else 1,
        help="10000 words級の論文では、まず段落単位がおすすめです。",
    )
    page_size = st.selectbox("1ページの表示件数", [10, 20, 50, 100, 200], index=1)
    context_radius = st.slider("前後文脈の表示範囲", min_value=1, max_value=5, value=2)
    highlight_height = st.slider("ハイライトビューの高さ", min_value=360, max_value=900, value=560, step=40)

    st.markdown("---")
    st.markdown("### 方針")
    st.markdown("- AIモデルなしで軽量")
    st.markdown("- 候補はドロップダウン選択")
    st.markdown("- 登録時に重複チェック")
    st.markdown("- 作業ページを切り替えて整理")

# 読み込み・新規作成
setup_left, setup_right = st.columns(2)

with setup_left:
    with st.expander("📂 保存済みJSONから作業を再開", expanded=False):
        uploaded_json = st.file_uploader("JSONファイルを選択してください", type=["json"])
        if uploaded_json is not None:
            st.write(f"選択中：{uploaded_json.name}")
            if st.button("このJSONを読み込んで再開", type="primary"):
                try:
                    title, unit, loaded_rows = load_rows_from_json(uploaded_json)
                    st.session_state["project_title"] = title
                    st.session_state["unit"] = unit if unit in ["段落", "文"] else "段落"
                    st.session_state["selected_id"] = loaded_rows[0]["ID"] if loaded_rows else None
                    st.success(f"JSONを読み込みました。{len(loaded_rows)}件のペアを復元しました。")
                    st.rerun()
                except Exception as error:
                    st.error(f"JSONの読み込みに失敗しました：{error}")

with setup_right:
    if st.session_state["rows"]:
        st.markdown(
            """
<div class="ppr-card">
  <b>現在の作業データ</b><br>
  <span class="ppr-subtle">既存プロジェクトを保持しています。新規作成すると本文ペア・マーカー・リンクはリセットされます。</span>
</div>
""",
            unsafe_allow_html=True,
        )

with st.expander("📝 テキストから新規作成・章を追加", expanded=(len(st.session_state["rows"]) == 0)):
    st.info("英文と和訳を同じ単位で貼り付けてください。自動対応づけはせず、段落または文の順番でペアを作ります。")

    section_title = st.text_input("章・節タイトル", value="Section 1")
    sample_english = """This study examines how students read academic papers.

Previous research has shown that translation tools can reduce reading difficulty.

However, students often struggle to compare the original English text with the Japanese translation."""
    sample_japanese = """本研究では、学生が学術論文をどのように読むかを検討する。

先行研究では、翻訳ツールが読解の困難を軽減できることが示されている。

しかし、学生は元の英文と日本語訳を比較する際に困難を感じることが多い。"""

    input_left, input_right = st.columns(2)
    with input_left:
        english_text = st.text_area("英文", value=sample_english if len(st.session_state["rows"]) == 0 else "", height=240)
    with input_right:
        japanese_text = st.text_area("和訳", value=sample_japanese if len(st.session_state["rows"]) == 0 else "", height=240)

    english_segments = segment_text(english_text, "英語", st.session_state["unit"])
    japanese_segments = segment_text(japanese_text, "日本語", st.session_state["unit"])
    english_word_count = estimate_english_words(english_text)

    metric1, metric2, metric3 = st.columns(3)
    metric1.metric("英文セグメント", len(english_segments))
    metric2.metric("和訳セグメント", len(japanese_segments))
    metric3.metric("英語語数の目安", english_word_count)

    if english_word_count >= 10000:
        st.warning("10000 words級です。章ごと・段落単位で進めるのがおすすめです。")
    if len(english_segments) != len(japanese_segments):
        st.warning("英文と和訳のセグメント数が違います。足りない側は空欄ペアになります。")

    create_col, append_col = st.columns(2)
    with create_col:
        if st.button("新規プロジェクトとして作成", type="primary"):
            if not english_segments and not japanese_segments:
                st.error("英文または和訳を入力してください。")
            else:
                new_rows = make_rows_from_segments(english_segments, japanese_segments, section_title, start_id=1)
                st.session_state["rows"] = new_rows
                st.session_state["markers"] = []
                st.session_state["links"] = []
                st.session_state["selected_id"] = 1 if new_rows else None
                cleanup_state()
                st.success(f"新規プロジェクトを作成しました。{len(new_rows)}件のペアがあります。")
                st.rerun()

    with append_col:
        if st.button("現在のプロジェクトに章として追加"):
            if not english_segments and not japanese_segments:
                st.error("英文または和訳を入力してください。")
            else:
                start_id = len(st.session_state["rows"]) + 1
                added = make_rows_from_segments(english_segments, japanese_segments, section_title, start_id=start_id)
                st.session_state["rows"].extend(added)
                st.session_state["rows"] = renumber_rows(st.session_state["rows"])
                st.session_state["selected_id"] = start_id
                cleanup_state()
                st.success(f"{len(added)}件を追加しました。")
                st.rerun()

rows = st.session_state["rows"]
if not rows:
    st.markdown("---")
    st.subheader("使い方")
    st.write("1. 英文と和訳を貼ります。")
    st.write("2. 10000 words級なら、処理単位はまず『段落』にします。")
    st.write("3. 『新規プロジェクトとして作成』を押します。")
    st.write("4. 上部の作業ページで、一覧選択・修正・マーカー・英日リンク・用語辞書・保存を行います。")
    st.stop()

# 概要
status_counts = {status: 0 for status in STATUS_OPTIONS}
for row in rows:
    status_counts[normalize_status(row.get("状態"))] += 1
marked_ids = marked_row_ids()
linked_ids = linked_row_ids()

cols = st.columns(7)
cols[0].metric("全ペア", len(rows))
cols[1].metric("未確認", status_counts["未確認"])
cols[2].metric("確認済み", status_counts["確認済み"])
cols[3].metric("要修正", status_counts["要修正"])
cols[4].metric("修正済み", status_counts["修正済み"])
cols[5].metric("マーカー行", len(marked_ids))
cols[6].metric("英日リンク行", len(linked_ids))

# ページナビUI
# st.tabs は全タブを毎回レンダリングするため、候補選択時に表示が崩れたり重くなったりしやすい。
# radio によるページ切り替えにして、選択中ページだけを描画する。
PAGES = ["一覧・選択", "読む・修正", "マーカー・英日リンク", "用語辞書", "保存"]
if st.session_state.get("active_page") not in PAGES:
    st.session_state["active_page"] = "一覧・選択"
if st.session_state.get("selected_id") is None or get_row_by_id(rows, st.session_state.get("selected_id")) is None:
    st.session_state["selected_id"] = int(rows[0]["ID"])

# Streamlit は、widget 生成後に同じ key の session_state を書き換えると例外を出す。
# そのため、通常のページ選択は radio widget に任せる。
# ボタンによるページ移動だけ、radio 生成前に pending 値として反映する。
pending_page = st.session_state.get("_pending_active_page")
if pending_page in PAGES:
    st.session_state["_active_page_widget"] = pending_page
    st.session_state["active_page"] = pending_page
    st.session_state["_pending_active_page"] = None
elif st.session_state.get("_active_page_widget") not in PAGES:
    st.session_state["_active_page_widget"] = st.session_state.get("active_page", "一覧・選択")

active_page = st.radio(
    "作業ページ",
    PAGES,
    horizontal=True,
    key="_active_page_widget",
    label_visibility="collapsed",
)

st.session_state["active_page"] = active_page
st.caption("ページ方式にしたため、候補選択時のタブ崩れを避け、必要な画面だけを描画します。")
st.markdown("---")

if active_page == "一覧・選択":
    st.markdown("### 一覧・選択")
    st.caption("行をクリックすると、他の作業ページの対象ペアが切り替わります。")

    all_sections = list(dict.fromkeys([str(row.get("章", "Section")) for row in rows]))
    f1, f2, f3 = st.columns([1.3, 1.1, 1.4])
    with f1:
        selected_sections = st.multiselect("章", options=all_sections, default=all_sections)
    with f2:
        selected_statuses = st.multiselect("状態", options=STATUS_OPTIONS, default=STATUS_OPTIONS)
    with f3:
        keyword = st.text_input("検索", value="")

    ff1, ff2, ff3, ff4 = st.columns(4)
    with ff1:
        only_marked = st.checkbox("マーカーあり")
    with ff2:
        only_linked = st.checkbox("英日リンクあり")
    with ff3:
        only_glossary = st.checkbox("用語辞書ヒット")
    with ff4:
        if st.button("重複登録を整理"):
            cleanup_state()
            st.success("重複登録を整理しました。")
            st.rerun()

    filtered_rows = filter_rows(rows, selected_sections, selected_statuses, keyword, only_marked, only_linked, only_glossary)
    st.session_state["last_filtered_ids"] = [int(row["ID"]) for row in filtered_rows]
    if not filtered_rows:
        st.warning("条件に合うペアがありません。")
        st.stop()

    total_pages = max(1, math.ceil(len(filtered_rows) / page_size))
    page = st.number_input("ページ", min_value=1, max_value=total_pages, value=1, step=1)
    start = (page - 1) * page_size
    page_rows = filtered_rows[start:start + page_size]
    st.caption(f"表示：{start + 1}〜{min(start + page_size, len(filtered_rows))}件 / フィルター後 {len(filtered_rows)}件 / 全体 {len(rows)}件")

    df = overview_dataframe(page_rows)
    event = st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        on_select="rerun",
        selection_mode="single-row",
    )
    selected_rows = event.selection.rows if hasattr(event, "selection") else []
    if selected_rows:
        st.session_state["selected_id"] = int(df.iloc[selected_rows[0]]["ID"])

    current_ids = [int(row["ID"]) for row in filtered_rows]
    if st.session_state.get("selected_id") not in current_ids:
        st.session_state["selected_id"] = current_ids[0]

    st.session_state["selected_id"] = int(
        st.selectbox(
            "編集するペアID",
            options=current_ids,
            index=current_ids.index(st.session_state["selected_id"]),
        )
    )
    jump_col1, jump_col2 = st.columns([1, 3])
    with jump_col1:
        if st.button("読む・修正へ"):
            st.session_state["_pending_active_page"] = "読む・修正"
            st.rerun()
    with jump_col2:
        st.caption("一覧で選んだIDは、他ページでもそのまま対象になります。")

selected_row = get_row_by_id(rows, st.session_state["selected_id"])
if selected_row is None:
    st.error("選択中のペアが見つかりませんでした。")
    st.stop()

if active_page == "読む・修正":
    st.markdown("### 読む・修正")
    top_left, top_right = st.columns([1, 1])
    with top_left:
        st.markdown(
            f"""
<div class="ppr-card">
  <b>選択中：</b>ID {selected_row['ID']} ／ {html.escape(str(selected_row.get('章','')))}-{selected_row.get('章内ID')}<br>
  <span class="ppr-chip ppr-chip-blue">状態：{html.escape(str(selected_row.get('状態','')))}</span>
  <span class="ppr-chip ppr-chip-yellow">マーカー：{len(markers_for(selected_row['ID']))}</span>
  <span class="ppr-chip ppr-chip-blue">英日リンク：{len(links_for(selected_row['ID']))}</span>
</div>
""",
            unsafe_allow_html=True,
        )
    with top_right:
        n1, n2, n3 = st.columns(3)
        current_index = get_row_index_by_id(rows, selected_row["ID"])
        with n1:
            if st.button("← 前", disabled=(current_index is None or current_index <= 0)):
                st.session_state["selected_id"] = rows[current_index - 1]["ID"]
                st.rerun()
        with n2:
            if st.button("空ペア追加"):
                insert_at = current_index + 1
                rows.insert(
                    insert_at,
                    {
                        "ID": 0,
                        "章": selected_row.get("章", "Section"),
                        "章内ID": 0,
                        "英文": "",
                        "原訳": "",
                        "修正訳": "",
                        "状態": "未確認",
                        "メモ": "",
                        "更新日時": now_string(),
                    },
                )
                st.session_state["rows"] = renumber_rows(rows)
                st.session_state["selected_id"] = insert_at + 1
                cleanup_state()
                st.rerun()
        with n3:
            if st.button("次 →", disabled=(current_index is None or current_index >= len(rows) - 1)):
                st.session_state["selected_id"] = rows[current_index + 1]["ID"]
                st.rerun()

    left, right = st.columns([1.2, 1])
    with left:
        components.html(
            build_highlight_html(selected_row, highlight_height),
            height=highlight_height + 74,
            scrolling=False,
        )
        with st.expander("前後文脈を見る", expanded=False):
            render_context(rows, selected_row["ID"], context_radius)

    with right:
        with st.form(key=f"edit_form_{selected_row['ID']}"):
            st.text_area("英文（確認用）", value=str(selected_row.get("英文", "")), height=170, disabled=True)
            edited_translation = st.text_area("修正訳", value=get_final_translation(selected_row), height=230)
            status = st.selectbox(
                "状態",
                STATUS_OPTIONS,
                index=STATUS_OPTIONS.index(normalize_status(selected_row.get("状態"))),
            )
            memo = st.text_area("メモ", value=str(selected_row.get("メモ", "")), height=80)
            restore = st.checkbox("原訳に戻す")
            submitted = st.form_submit_button("このペアを保存", type="primary")

        if submitted:
            row_index = get_row_index_by_id(rows, selected_row["ID"])
            if row_index is not None:
                rows[row_index]["修正訳"] = str(selected_row.get("原訳", "")) if restore else edited_translation
                rows[row_index]["状態"] = status
                rows[row_index]["メモ"] = memo
                rows[row_index]["更新日時"] = now_string()
                st.session_state["rows"] = rows
                st.success("保存しました。")
                st.rerun()

        with st.expander("結合・削除", expanded=False):
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("次の訳文を結合", disabled=(current_index is None or current_index >= len(rows) - 1)):
                    current = rows[current_index]
                    next_row = rows[current_index + 1]
                    next_id = int(next_row["ID"])
                    current["原訳"] = "\n\n".join([p for p in [str(current.get("原訳", "")).strip(), str(next_row.get("原訳", "")).strip()] if p])
                    current["修正訳"] = "\n\n".join([p for p in [get_final_translation(current), get_final_translation(next_row)] if p])
                    current["状態"] = "修正済み"
                    for m in st.session_state["markers"]:
                        if int(m["row_id"]) == next_id:
                            m["row_id"] = int(current["ID"])
                    for l in st.session_state["links"]:
                        if int(l["row_id"]) == next_id:
                            l["row_id"] = int(current["ID"])
                    del rows[current_index + 1]
                    st.session_state["rows"] = renumber_rows(rows)
                    cleanup_state()
                    st.rerun()
            with c2:
                if st.button("次の英文を結合", disabled=(current_index is None or current_index >= len(rows) - 1)):
                    current = rows[current_index]
                    next_row = rows[current_index + 1]
                    next_id = int(next_row["ID"])
                    current["英文"] = "\n\n".join([p for p in [str(current.get("英文", "")).strip(), str(next_row.get("英文", "")).strip()] if p])
                    current["状態"] = "修正済み"
                    for m in st.session_state["markers"]:
                        if int(m["row_id"]) == next_id:
                            m["row_id"] = int(current["ID"])
                    for l in st.session_state["links"]:
                        if int(l["row_id"]) == next_id:
                            l["row_id"] = int(current["ID"])
                    del rows[current_index + 1]
                    st.session_state["rows"] = renumber_rows(rows)
                    cleanup_state()
                    st.rerun()
            with c3:
                delete_ok = st.checkbox("削除許可")
                if st.button("削除", disabled=(not delete_ok or len(rows) <= 1)):
                    deleted_id = int(selected_row["ID"])
                    del rows[current_index]
                    st.session_state["markers"] = [m for m in st.session_state["markers"] if int(m["row_id"]) != deleted_id]
                    st.session_state["links"] = [l for l in st.session_state["links"] if int(l["row_id"]) != deleted_id]
                    st.session_state["rows"] = renumber_rows(rows)
                    cleanup_state()
                    st.session_state["selected_id"] = min(deleted_id, len(rows))
                    st.rerun()

if active_page == "マーカー・英日リンク":
    st.markdown("### マーカー・英日リンク")
    st.caption("候補はボタンではなくドロップダウンにしました。選んで追加すると、重複なしで登録されます。")

    marker_col, link_col = st.columns(2)

    with marker_col:
        st.markdown("#### 日本語マーカー")
        ja_candidates = [""] + suggest_japanese_phrases(get_final_translation(selected_row))
        with st.form(key=f"marker_form_{selected_row['ID']}"):
            selected_candidate = st.selectbox("候補から選ぶ", ja_candidates, format_func=lambda x: "候補を選択" if x == "" else x)
            manual_phrase = st.text_input("または日本語句を直接入力", placeholder="例：読解の困難")
            marker_label = st.selectbox("ラベル", MARKER_LABELS)
            marker_note = st.text_input("メモ", placeholder="例：意味が少し曖昧")
            add_marker = st.form_submit_button("マーカー追加", type="primary")
        if add_marker:
            phrase = manual_phrase.strip() if manual_phrase.strip() else selected_candidate
            ok, message = add_unique_marker(selected_row["ID"], phrase, marker_label, marker_note)
            if ok:
                st.success(message)
                st.rerun()
            else:
                st.warning(message)

        current_markers = markers_for(selected_row["ID"])
        if current_markers:
            st.dataframe(
                pd.DataFrame(
                    [
                        {"id": m["id"], "ラベル": m["label"], "日本語句": m["ja_phrase"], "メモ": m["note"]}
                        for m in current_markers
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
            delete_marker_id = st.selectbox("削除するマーカーID", [int(m["id"]) for m in current_markers], key=f"delete_marker_{selected_row['ID']}")
            if st.button("選択マーカーを削除"):
                st.session_state["markers"] = [m for m in st.session_state["markers"] if int(m["id"]) != int(delete_marker_id)]
                cleanup_state()
                st.rerun()
        else:
            st.info("このペアには、まだマーカーがありません。")

    with link_col:
        st.markdown("#### 英日リンク")
        ja_link_candidates = [""] + suggest_japanese_phrases(get_final_translation(selected_row))
        en_link_candidates = [""] + suggest_english_phrases(selected_row.get("英文", ""))
        with st.form(key=f"link_form_{selected_row['ID']}"):
            selected_ja = st.selectbox("日本語候補", ja_link_candidates, format_func=lambda x: "候補を選択" if x == "" else x)
            selected_en = st.selectbox("英語候補", en_link_candidates, format_func=lambda x: "候補を選択" if x == "" else x)
            manual_ja = st.text_input("または日本語句を直接入力", placeholder="例：読解の困難")
            manual_en = st.text_input("または英語句を直接入力", placeholder="例：reading difficulty")
            link_label = st.selectbox("リンク種別", LINK_LABELS)
            link_note = st.text_input("メモ", placeholder="例：訳語確認")
            add_link = st.form_submit_button("英日リンク追加", type="primary")
        if add_link:
            ja_phrase = manual_ja.strip() if manual_ja.strip() else selected_ja
            en_phrase = manual_en.strip() if manual_en.strip() else selected_en
            ok, message = add_unique_link(selected_row["ID"], ja_phrase, en_phrase, link_label, link_note)
            if ok:
                st.success(message)
                st.rerun()
            else:
                st.warning(message)

        current_links = links_for(selected_row["ID"])
        if current_links:
            st.dataframe(
                pd.DataFrame(
                    [
                        {"id": l["id"], "種別": l["label"], "日本語句": l["ja_phrase"], "英語句": l["en_phrase"], "メモ": l["note"]}
                        for l in current_links
                    ]
                ),
                use_container_width=True,
                hide_index=True,
            )
            delete_link_id = st.selectbox("削除する英日リンクID", [int(l["id"]) for l in current_links], key=f"delete_link_{selected_row['ID']}")
            if st.button("選択リンクを削除"):
                st.session_state["links"] = [l for l in st.session_state["links"] if int(l["id"]) != int(delete_link_id)]
                cleanup_state()
                st.rerun()
        else:
            st.info("このペアには、まだ英日リンクがありません。")

    with st.expander("全マーカー・全リンクを一覧で見る", expanded=False):
        all1, all2 = st.columns(2)
        with all1:
            st.markdown("**全マーカー**")
            if st.session_state["markers"]:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "id": m["id"],
                                "ID": m["row_id"],
                                "ラベル": m["label"],
                                "日本語句": m["ja_phrase"],
                                "メモ": m["note"],
                            }
                            for m in st.session_state["markers"]
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("まだありません。")
        with all2:
            st.markdown("**全英日リンク**")
            if st.session_state["links"]:
                st.dataframe(
                    pd.DataFrame(
                        [
                            {
                                "id": l["id"],
                                "ID": l["row_id"],
                                "種別": l["label"],
                                "日本語句": l["ja_phrase"],
                                "英語句": l["en_phrase"],
                                "メモ": l["note"],
                            }
                            for l in st.session_state["links"]
                        ]
                    ),
                    use_container_width=True,
                    hide_index=True,
                )
            else:
                st.info("まだありません。")

if active_page == "用語辞書":
    st.markdown("### 用語辞書")
    st.caption("登録した語は、選択中ペアのハイライトビューで緑色になります。")

    g_left, g_right = st.columns([1, 1.3])
    with g_left:
        with st.form("glossary_form"):
            glossary_en = st.text_input("英語用語", placeholder="例：self-regulated learning")
            glossary_ja = st.text_input("日本語訳", placeholder="例：自己調整学習")
            glossary_category = st.selectbox("カテゴリ", GLOSSARY_CATEGORIES)
            glossary_note = st.text_input("メモ", placeholder="訳語の使い分けなど")
            add_glossary = st.form_submit_button("用語を追加", type="primary")
        if add_glossary:
            ok, message = add_unique_glossary(glossary_en, glossary_ja, glossary_category, glossary_note)
            if ok:
                st.success(message)
                st.rerun()
            else:
                st.warning(message)

    with g_right:
        if st.session_state["glossary"]:
            glossary_df = pd.DataFrame(
                [
                    {
                        "id": item["id"],
                        "英語": item["english"],
                        "日本語": item["japanese"],
                        "カテゴリ": item["category"],
                        "メモ": item["note"],
                    }
                    for item in st.session_state["glossary"]
                ]
            )
            st.dataframe(glossary_df, use_container_width=True, hide_index=True)
            delete_glossary_id = st.selectbox("削除する用語ID", [int(g["id"]) for g in st.session_state["glossary"]])
            if st.button("選択用語を削除"):
                st.session_state["glossary"] = [g for g in st.session_state["glossary"] if int(g["id"]) != int(delete_glossary_id)]
                cleanup_state()
                st.rerun()
        else:
            st.info("まだ用語辞書は空です。")

if active_page == "保存":
    st.markdown("### 保存・Word出力")
    st.caption("出力データは、ボタンを押した時だけ作成します。大きい論文でも再描画のたびにWordを作らないので軽くなります。")

    save_target = st.radio("保存対象", ["全件", "フィルター後のみ"], horizontal=True)
    if save_target == "全件":
        export_rows = rows
    else:
        last_filtered_ids = [int(x) for x in st.session_state.get("last_filtered_ids", [])]
        if last_filtered_ids:
            id_set = set(last_filtered_ids)
            export_rows = [row for row in rows if int(row["ID"]) in id_set]
        else:
            export_rows = rows
        st.caption(f"フィルター後のみ：{len(export_rows)}件（最後に『一覧・選択』で適用した条件を使用）")

    export_style = st.selectbox("Word出力形式", WORD_EXPORT_OPTIONS)
    base_name = safe_filename(st.session_state["project_title"])

    st.markdown(
        f"""
<div class="ppr-card">
  <b>出力予定：</b>{len(export_rows)}件<br>
  <span class="ppr-subtle">修正後は、もう一度「出力ファイルを準備」を押すと最新内容で作り直します。</span>
</div>
""",
        unsafe_allow_html=True,
    )

    if st.button("出力ファイルを準備", type="primary"):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        st.session_state["export_cache"] = {
            "base_name": base_name,
            "timestamp": timestamp,
            "row_count": len(export_rows),
            "csv": make_csv_data(export_rows),
            "json": make_json_data(st.session_state["project_title"], st.session_state["unit"], rows),
            "docx": make_docx_data(st.session_state["project_title"], export_rows, export_style),
        }
        st.success("出力ファイルを準備しました。下のボタンから保存できます。")

    cache = st.session_state.get("export_cache", {})
    if cache:
        prepared_name = cache.get("base_name", base_name)
        prepared_timestamp = cache.get("timestamp", datetime.now().strftime("%Y%m%d_%H%M%S"))
        save_left, save_center, save_right = st.columns(3)
        with save_left:
            st.download_button(
                "CSVで保存",
                data=cache["csv"],
                file_name=f"{prepared_name}_{prepared_timestamp}.csv",
                mime="text/csv",
            )
        with save_center:
            st.download_button(
                "JSONで保存・再開用",
                data=cache["json"],
                file_name=f"{prepared_name}_{prepared_timestamp}.json",
                mime="application/json",
            )
        with save_right:
            st.download_button(
                "Wordで保存",
                data=cache["docx"],
                file_name=f"{prepared_name}_{prepared_timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    else:
        st.info("まず『出力ファイルを準備』を押してください。JSONには、本文ペア・修正訳・状態・メモ・用語辞書・日本語マーカー・英日リンクがすべて保存されます。")
