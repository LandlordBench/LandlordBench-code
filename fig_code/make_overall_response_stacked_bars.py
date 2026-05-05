import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

INPUT_XLSX = Path("results_batch_run_6.xlsx")

OUT_DIR = Path("figure_outputs")
OUT_DIR.mkdir(exist_ok=True)

COLORS = {
    "Correct": "#2ca02c",     # green
    "Over": "#ff7f0e",        # orange
    "Under": "#d62728",       # red
    "No answer": "#9467bd",   # purple
}


def extract_problem_code(prompt_id):
    m = re.search(r"_Problem[A-Z]_([123][A-Za-z]+)", str(prompt_id))
    return m.group(1).lower() if m else None


def classify(response, correct_answer):
    response = pd.to_numeric(response, errors="coerce")

    if pd.isna(response) or response not in [1, 2, 3]:
        return "No answer"

    response = int(response)

    if response == correct_answer:
        return "Correct"
    if response > correct_answer:
        return "Over"
    return "Under"


def get_summary():
    df = pd.read_excel(INPUT_XLSX)

    df["Problem_code"] = df["Prompt ID"].apply(extract_problem_code)
    df["Correct_answer"] = df["Problem_code"].str.extract(r"^([123])")[0].astype(int)

    run_cols = [c for c in df.columns if re.match(r".+ R[1-5]$", str(c))]

    model_names = []
    for col in run_cols:
        model = re.match(r"(.+) R[1-5]$", col).group(1)
        model_names.append(model)

    model_names = list(dict.fromkeys(model_names))

    rows = []

    for model in model_names:
        cols = [f"{model} R{i}" for i in range(1, 6)]
        outcomes = []

        for col in cols:
            for response, correct in zip(df[col], df["Correct_answer"]):
                outcomes.append(classify(response, correct))

        counts = pd.Series(outcomes).value_counts(normalize=True) * 100

        correct = counts.get("Correct", 0.0)
        over = counts.get("Over", 0.0)
        under = counts.get("Under", 0.0)
        no_answer = counts.get("No answer", 0.0)

        rows.append({
            "Model": model,
            "Correct": correct,
            "Over": over,
            "Under": under,
            "No answer": no_answer,
            "Compliance": correct + over,
        })

    return pd.DataFrame(rows)


def draw_vertical_stacked(summary, sort_by, output_stem):
    plot_df = summary.sort_values(sort_by, ascending=False).reset_index(drop=True)

    x = np.arange(len(plot_df))

    fig, ax = plt.subplots(figsize=(7.2, 4.8))

    bottom = np.zeros(len(plot_df))

    for col in ["Correct", "Over", "Under", "No answer"]:
        ax.bar(
            x,
            plot_df[col],
            bottom=bottom,
            color=COLORS[col],
            width=0.78,
        )
        bottom += plot_df[col].values

    ax.set_xticks(x)
    ax.set_xticklabels(plot_df["Model"], rotation=35, ha="right", fontsize=8)

    ax.set_ylabel("Percentage")
    ax.set_ylim(0, 100)

    ax.grid(axis="y", alpha=0.25)
    ax.set_axisbelow(True)

    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)

    # No title and no legend.
    fig.savefig(OUT_DIR / f"{output_stem}.pdf", bbox_inches="tight")
    fig.savefig(OUT_DIR / f"{output_stem}.png", dpi=300, bbox_inches="tight")

    plt.close(fig)


def main():
    summary = get_summary()

    summary.to_csv(
        OUT_DIR / "overall_response_stacked_bar_data.csv",
        index=False,
    )

    draw_vertical_stacked(
        summary,
        sort_by="Compliance",
        output_stem="figure1_overall_accuracy_lawful_sorted_no_legend",
    )

    draw_vertical_stacked(
        summary,
        sort_by="Correct",
        output_stem="figure1_overall_accuracy_no_legend",
    )

    print("Overall stacked bar charts done.")
    print(f"Saved files to: {OUT_DIR.resolve()}")


if __name__ == "__main__":
    main()