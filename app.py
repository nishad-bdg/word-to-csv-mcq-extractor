import re
import io
import csv
import random
import json
from typing import List, Dict, Any

import streamlit as st
import pandas as pd
from docx import Document

# SAME COLUMNS AS YOUR TEMPLATE (do not change order)
COLUMNS = [
    "question_number",
    "question",
    "subject",
    "paper",
    "chapter",
    "topic",
    "source",
    "tags",
    "choice_A",
    "choice_B",
    "choice_C",
    "choice_D",
    "correct_answer",
    "correct_index",
    "correct_text",
    "answer_explanation",
]

TAG_POOL = [
    "bangla", "mcq", "exam", "language", "grammar",
    "synonym", "practice", "board", "modeltest"
]

LABEL_TO_INDEX = {"ক": 0, "খ": 1, "গ": 2, "ঘ": 3}
INDEX_TO_COL = {0: "choice_A", 1: "choice_B", 2: "choice_C", 3: "choice_D"}


def docx_bytes_to_text(file_bytes: bytes) -> str:
    doc = Document(io.BytesIO(file_bytes))
    return "\n".join(p.text.strip() for p in doc.paragraphs if p.text and p.text.strip())


def normalize_spaces(s: str) -> str:
    return " ".join(s.split()).strip()


def parse_mcqs(text: str, default_subject: str, default_paper: str) -> List[Dict[str, Any]]:
    """
    Parses MCQs that look like:

    1. Question text...
    ক) option
    খ) option
    গ) option
    ঘ) option
    উত্তর: খ
    """

    # Split by lines starting with "1.", "2.", etc.
    blocks = re.split(r"\n(?=\d+\.)", text)
    items: List[Dict[str, Any]] = []

    for block in blocks:
        block = block.strip()
        if not block:
            continue

        qno_match = re.match(r"(\d+)\.\s*(.*)", block, re.S)
        if not qno_match:
            continue

        qno = qno_match.group(1).zfill(2)
        body = qno_match.group(2).strip()

        # Find the first option marker (ক)
        first_opt = re.search(r"\n\s*ক\s*[.)\-ঃ:]\s*", body)
        if not first_opt:
            continue

        question = normalize_spaces(body[: first_opt.start()])
        opt_text = body[first_opt.start():].strip()

        # Extract options: (label, text until next label or end)
        opt_matches = re.findall(
            r"(ক|খ|গ|ঘ)\s*[.)\-ঃ:]\s*(.*?)(?=\n\s*(?:ক|খ|গ|ঘ)\s*[.)\-ঃ:]|\Z)",
            opt_text,
            flags=re.S,
        )

        options = ["", "", "", ""]
        for lbl, txt in opt_matches:
            idx = LABEL_TO_INDEX.get(lbl)
            if idx is not None:
                options[idx] = normalize_spaces(txt)

        # Detect correct answer
        # supports: "উত্তর: খ" / "উ. খ" / "সঠিক উত্তর: খ"
        ans_match = re.search(
            r"(?:উত্তর|উ\.|সঠিক\s*উত্তর)\s*[:：\-]?\s*(ক|খ|গ|ঘ)",
            block
        )
        correct_label = ans_match.group(1) if ans_match else ""

        # ✅ Make correct_index ALWAYS STRING for Arrow safety
        if correct_label in LABEL_TO_INDEX:
            correct_index_str = str(LABEL_TO_INDEX[correct_label])
            correct_text = options[int(correct_index_str)]
        else:
            correct_index_str = ""   # always string
            correct_text = ""

        # Random tags as comma-separated
        tags = ",".join(random.sample(TAG_POOL, k=3)) if len(TAG_POOL) >= 3 else ""

        row = {col: "" for col in COLUMNS}
        row.update({
            "question_number": qno,
            "question": question,
            "subject": default_subject.strip(),
            "paper": default_paper.strip(),
            "chapter": "",
            "topic": "",
            "source": "",
            "tags": tags,
            "choice_A": options[0],
            "choice_B": options[1],
            "choice_C": options[2],
            "choice_D": options[3],
            "correct_answer": correct_label,
            "correct_index": correct_index_str,  # ✅ string
            "correct_text": correct_text,
            "answer_explanation": f"সঠিক উত্তর: {correct_label}" if correct_label else "",
        })

        # Ensure required fields
        if row["subject"] and row["question"]:
            items.append(row)

    return items


def rows_to_csv_bytes(rows: List[Dict[str, Any]]) -> bytes:
    buf = io.StringIO()
    writer = csv.DictWriter(buf, fieldnames=COLUMNS)
    writer.writeheader()
    for r in rows:
        # Ensure all columns exist + correct_index is string
        safe_row = {k: r.get(k, "") for k in COLUMNS}
        safe_row["correct_index"] = "" if safe_row["correct_index"] is None else str(safe_row["correct_index"])
        writer.writerow(safe_row)
    return buf.getvalue().encode("utf-8-sig")


def rows_to_json_bytes(rows: List[Dict[str, Any]]) -> bytes:
    # Ensure correct_index is string in JSON too (optional but consistent)
    fixed = []
    for r in rows:
        rr = dict(r)
        rr["correct_index"] = "" if rr.get("correct_index") in (None, "") else str(rr.get("correct_index"))
        fixed.append(rr)
    return json.dumps(fixed, ensure_ascii=False, indent=2).encode("utf-8")


def rows_to_safe_dataframe(rows: List[Dict[str, Any]]) -> pd.DataFrame:
    """
    Streamlit uses Arrow for dataframe display.
    Arrow crashes when dtype is object with mixed int/str.
    So we force specific columns to string.
    """
    df = pd.DataFrame(rows)

    # If empty, return as-is
    if df.empty:
        return df

    # Force these columns to string (safe for Arrow)
    force_string_cols = [
        "question_number",
        "subject",
        "paper",
        "chapter",
        "topic",
        "source",
        "tags",
        "choice_A",
        "choice_B",
        "choice_C",
        "choice_D",
        "correct_answer",
        "correct_index",
        "correct_text",
        "answer_explanation",
        "question",
    ]
    for c in force_string_cols:
        if c in df.columns:
            df[c] = df[c].fillna("").astype(str)

    return df


# ------------------- UI -------------------
st.set_page_config(page_title="DOCX → MCQ CSV Converter", layout="wide")

st.title("📄➡️📊 DOCX → MCQ CSV/JSON Converter")
st.write("Upload a Word file, extract MCQs, and download in your **exact CSV template format** (all columns included).")

col1, col2, col3 = st.columns(3)
with col1:
    default_subject = st.text_input("Default subject", value="বাংলা")
with col2:
    default_paper = st.text_input("Default paper", value="")
with col3:
    st.caption("Tags will be auto-filled with random values.")

uploaded = st.file_uploader("Upload .docx file", type=["docx"])

if uploaded:
    file_bytes = uploaded.read()
    raw_text = docx_bytes_to_text(file_bytes)

    with st.expander("Preview extracted text"):
        st.text(raw_text[:5000] + ("\n\n... (trimmed)" if len(raw_text) > 5000 else ""))

    rows = parse_mcqs(raw_text, default_subject=default_subject, default_paper=default_paper)

    st.subheader("✅ Extraction result")
    st.write(f"Found **{len(rows)}** MCQ(s).")

    if rows:
        # ✅ SAFE dataframe display (prevents Arrow crash)
        df = rows_to_safe_dataframe(rows)
        st.dataframe(df, use_container_width=True)

        csv_bytes = rows_to_csv_bytes(rows)
        json_bytes = rows_to_json_bytes(rows)

        c1, c2 = st.columns(2)
        with c1:
            st.download_button(
                label="⬇️ Download CSV",
                data=csv_bytes,
                file_name="mcq_output.csv",
                mime="text/csv",
            )
        with c2:
            st.download_button(
                label="⬇️ Download JSON",
                data=json_bytes,
                file_name="mcq_output.json",
                mime="application/json",
            )
    else:
        st.error("No MCQs detected. Check if the file has numbered questions and options like ক/খ/গ/ঘ.")
