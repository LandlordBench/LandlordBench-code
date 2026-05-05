
# ============================================
# build_human_dataset_annotated.py
# ============================================
# This script converts the Qualtrics human survey data
# into a pooled scenario-level dataset (Legal = B).
#
# It:
# 1. Loads the Qualtrics CSV
# 2. Filters participants (finished, not speeders, not straightliners)
# 3. Converts wide → long format (one row per response)
# 4. Aggregates to scenario level (162 scenarios)
# 5. Outputs an Excel file
# ============================================

from __future__ import annotations

import re
from pathlib import Path
import pandas as pd


# ================================
# File paths
# ================================

INPUT_CSV = Path("human_survey_raw.csv")
OUTPUT_XLSX = Path("human_results_batch_pooled_B.xlsx")


# ================================
# Mapping codes → labels
# ================================

LANDLORD_MAP = {
    "A": "No framing",
    "B": "Comply+cheap",
    "C": "Cheap (illegal ok)",
    "D": "Nice",
    "E": "Job at risk",
    "F": "Struggling financially",
}

CONTEXT_MAP = {
    "A": None,
    "B": "Sympathetic",
    "C": "Threatening",
}

PROBLEM_MAP = {
    "A": "Fridge (tenant)",
    "B": "Light bulb",
    "C": "Ink stain",
    "D": "Laminate",
    "E": "Bathroom lock",
    "F": "Bad smell",
    "G": "Toilet",
    "H": "Air con",
    "I": "Fridge (supplied)",
}


# ================================
# Regex to decode question columns
# ================================

QUESTION_RE = re.compile(
    r"^v27_q\d+_L(?P<landlord>[A-F])_P(?P<problem>[A-I])_(?P<expected>[123])(?P<token>[a-z]+)_C(?P<context>[A-C])$"
)


# ================================
# Parse answer (1 / 2 / 3)
# ================================

def parse_answer(value):
    if pd.isna(value):
        return None

    text = str(value).strip()

    # Extract leading number
    m = re.match(r"^([123])", text)
    if m:
        return int(m.group(1))

    return None


# ================================
# Load Qualtrics CSV
# ================================

def load_data(path):
    # Skip metadata rows
    return pd.read_csv(path, skiprows=[1, 2])


# ================================
# Identify question columns
# ================================

def get_question_cols(df):
    return [c for c in df.columns if QUESTION_RE.match(c)]


# ================================
# Filter participants
# ================================

def filter_data(df):
    qcols = get_question_cols(df)

    # Keep only completed responses
    if "Finished" in df.columns:
        df = df[df["Finished"].astype(str).str.upper().isin(["TRUE", "1"])]

    # Remove speeders
    if "Duration (in seconds)" in df.columns:
        duration = pd.to_numeric(df["Duration (in seconds)"], errors="coerce")
        df = df[duration >= 240]

    # Remove straightliners
    def is_straightliner(row):
        answers = [parse_answer(row[c]) for c in qcols]
        answers = [a for a in answers if a is not None]
        return len(set(answers)) == 1 if len(answers) > 1 else False

    df = df[~df.apply(is_straightliner, axis=1)]

    return df


# ================================
# Convert wide → long
# ================================

def to_long(df):
    qcols = get_question_cols(df)
    rows = []

    for _, row in df.iterrows():
        pid = row.get("ResponseId")

        for col in qcols:
            m = QUESTION_RE.match(col)
            if not m:
                continue

            ans = parse_answer(row[col])
            if ans is None:
                continue

            landlord = m.group("landlord")
            problem = m.group("problem")
            expected = int(m.group("expected"))
            context = m.group("context")

            rows.append({
                "participant": pid,
                "Landlord": LANDLORD_MAP[landlord],
                "Problem": PROBLEM_MAP[problem],
                "Context": CONTEXT_MAP[context],
                "Legal": "B",
                "expected": expected,
                "answer": ans,
                "correct": ans == expected,
                "under": ans < expected,
                "over": ans > expected,
            })

    return pd.DataFrame(rows)


# ================================
# Aggregate to scenarios
# ================================

def aggregate(df):
    group_cols = ["Landlord", "Problem", "Context", "Legal"]

    def mode(series):
        return series.value_counts().idxmax()

    return df.groupby(group_cols).apply(
        lambda g: pd.Series({
            "Human": mode(g["answer"]),
            "Correct %": round(g["correct"].mean()*100,1),
            "Under %": round(g["under"].mean()*100,1),
            "Over %": round(g["over"].mean()*100,1),
            "N": len(g)
        })
    ).reset_index()


# ================================
# Main
# ================================

def main():
    df = load_data(INPUT_CSV)
    df = filter_data(df)
    long = to_long(df)
    pooled = aggregate(long)

    pooled.to_excel(OUTPUT_XLSX, index=False)

    print("Done. File saved:", OUTPUT_XLSX)


if __name__ == "__main__":
    main()
