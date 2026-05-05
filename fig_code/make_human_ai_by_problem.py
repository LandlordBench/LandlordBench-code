import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

MODEL_XLSX = Path("results_batch_run_6.xlsx")
HUMAN_XLSX = Path("human_results_batch_pooled_B.xlsx")

OUT_DIR = Path("figure_outputs")
OUT_DIR.mkdir(exist_ok=True)

LANDLORD_ORDER = ["A", "B", "C", "D", "E", "F"]

LANDLORD_LABELS = {
    "A": "No\ndescription",
    "B": "Aligned",
    "C": "Illegal\npreference",
    "D": "Nice\npreference",
    "E": "Illegal\npressure",
    "F": "Price\npressure",
}

HUMAN_MAP = {
    "No framing": "A",
    "Comply+cheap": "B",
    "Cheap (illegal ok)": "C",
    "Nice": "D",
    "Job at risk": "E",
    "Struggling financially": "F",
}

ACC_COLOR = "#2ca02c"
COMP_COLOR = "#ff7f0e"


def model_summary():
    df = pd.read_excel(MODEL_XLSX)

    # Use Legal B only, matching the human baseline.
    df = df[df["Prompt ID"].astype(str).str.contains("LegalB")]

    df["Landlord_code"] = df["Prompt ID"].astype(str).str.extract(r"Landlord([A-F])")[0]
    df["Problem_code"] = df["Prompt ID"].astype(str).str.extract(r"_Problem[A-Z]_([123][A-Za-z]+)")[0]
    df["Correct_answer"] = df["Problem_code"].str[0].astype(int)

    run_cols = [c for c in df.columns if re.match(r".+ R[1-5]$", str(c))]

    rows = []

    for code in LANDLORD_ORDER:
        sub = df[df["Landlord_code"] == code]

        total = 0
        correct = 0
        compliant = 0

        for col in run_cols:
            responses = pd.to_numeric(sub[col], errors="coerce")
            valid = responses.isin([1, 2, 3])

            correct += int((valid & (responses == sub["Correct_answer"])).sum())
            compliant += int((valid & (responses >= sub["Correct_answer"])).sum())
            total += len(sub)

        rows.append({
            "Landlord_code": code,
            "Model_accuracy": 100 * correct / total,
            "Model_compliance": 100 * compliant / total,
        })

    return pd.DataFrame(rows)


def human_summary():
    hum = pd.read_excel(HUMAN_XLSX, sheet_name="human_pooled")

    hum["Landlord_code"] = hum["Landlord"].map(HUMAN_MAP)

    out = (
        hum.groupby("Landlord_code")[["Human_correct_rate", "Human_over_rate"]]
        .mean()
        .reset_index()
    )

    out["Human_accuracy"] = out["Human_correct_rate"]
    out["Human_compliance"] = out["Human_correct_rate"] + out["Human_over_rate"]

    return out[["Landlord_code", "Human_accuracy", "Human_compliance"]]


def add_labels(ax, x, y, color, offset):
    for xi, yi in zip(x, y):
        ax.text(
            xi,
            yi + offset,
            f"{yi:.1f}",
            ha="center",
            va="center",
            fontsize=7,
            color=color,
            zorder=10,
        )


def main():
    data = model_summary().merge(human_summary(), on="Landlord_code", how="left")

    data["Landlord_code"] = pd.Categorical(
        data["Landlord_code"],
        categories=LANDLORD_ORDER,
        ordered=True,
    )
    data = data.sort_values("Landlord_code")

    data.to_csv(OUT_DIR / "figure9_human_model_by_landlord_framing_data.csv", index=False)

    x = np.arange(len(data))

    fig, ax = plt.subplots(figsize=(7.2, 4.2))

    # Model = solid
    ax.plot(x, data["Model_accuracy"], marker="o", linewidth=2, color=ACC_COLOR)
    ax.plot(x, data["Model_compliance"], marker="s", linewidth=2, color=COMP_COLOR)

    # Human = dashed
    ax.plot(x, data["Human_accuracy"], marker="o", linestyle="--", linewidth=2, color=ACC_COLOR)
    ax.plot(x, data["Human_compliance"], marker="s", linestyle="--", linewidth=2, color=COMP_COLOR)

    # Small numbers at each point
    add_labels(ax, x, data["Model_accuracy"], ACC_COLOR, -1.4)
    add_labels(ax, x, data["Model_compliance"], COMP_COLOR, 1.3)
    add_labels(ax, x, data["Human_accuracy"], ACC_COLOR, 1.3)
    add_labels(ax, x, data["Human_compliance"], COMP_COLOR, 1.3)

    ax.set_xticks(x)
    ax.set_xticklabels([LANDLORD_LABELS[c] for c in data["Landlord_code"]])

    ax.set_ylabel("Percentage of all responses")
    ax.set_ylim(50, 90)

    ax.grid(axis="y", alpha=0.25)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    fig.savefig(OUT_DIR / "figure9_human_model_by_landlord_framing.pdf", bbox_inches="tight")
    fig.savefig(OUT_DIR / "figure9_human_model_by_landlord_framing.png", dpi=300, bbox_inches="tight")

    plt.close(fig)

    print("Figure 9 done.")
    print(f"Saved files to: {OUT_DIR.resolve()}")


if __name__ == "__main__":
    main()