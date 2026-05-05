import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import re
from pathlib import Path

# ------------------------------------------------------------
# File paths
# ------------------------------------------------------------

INPUT_XLSX = "results_batch_run_6.xlsx"
OUT_DIR = Path("legal_condition_charts")
OUT_DIR.mkdir(exist_ok=True)

# ------------------------------------------------------------
# Settings
# ------------------------------------------------------------

LEGAL_ORDER = ["No law", "Excerpt", "Full schedule", "Schedule + cases"]

LEGAL_RENAME = {
    "No law": "No law",
    "Excerpt": "Excerpt",
    "Full schedule": "Full schedule",
    "Sched+relevant cases": "Schedule + cases",
}

# Same colour scale for both heatmaps
VMIN = 20
VMAX = 100

# ------------------------------------------------------------
# Load data
# ------------------------------------------------------------

df = pd.read_excel(INPUT_XLSX)

# Clean legal condition names
df["Legal_clean"] = df["Legal"].map(LEGAL_RENAME).fillna(df["Legal"])

# Extract correct answer from Prompt ID, e.g. "..._ProblemH_3aircon_..."
df["Correct answer"] = (
    df["Prompt ID"]
    .astype(str)
    .str.extract(r"_Problem[A-Z]_([123])")[0]
    .astype(int)
)

# Find model names from columns like "GPT-5.4 R1", "GPT-5.4 R2", etc.
model_names = []
for col in df.columns:
    match = re.match(r"(.+) R[1-5]$", col)
    if match:
        model_names.append(match.group(1))

model_names = list(dict.fromkeys(model_names))

# ------------------------------------------------------------
# Compute correct and compliant percentages
# ------------------------------------------------------------

rows = []

for model in model_names:
    run_cols = [f"{model} R{i}" for i in range(1, 6)]

    for legal in LEGAL_ORDER:
        sub = df[df["Legal_clean"] == legal]

        total = 0
        correct_count = 0
        compliant_count = 0

        for col in run_cols:
            responses = pd.to_numeric(sub[col], errors="coerce")
            correct_answer = sub["Correct answer"]

            # Valid answers are only 1, 2, or 3.
            # No answers are counted in the denominator but are not correct/compliant.
            valid = responses.isin([1, 2, 3])

            correct = valid & (responses == correct_answer)

            # Compliant = correct or over-recommendation.
            # Since answers are ordered by legal strictness/cost:
            # 1 = refuse, 2 = slow repair, 3 = fast repair.
            # Anything >= correct answer is lawful.
            compliant = valid & (responses >= correct_answer)

            total += len(sub)
            correct_count += correct.sum()
            compliant_count += compliant.sum()

        rows.append({
            "Model": model,
            "Legal": legal,
            "Correct": 100 * correct_count / total,
            "Compliant": 100 * compliant_count / total,
        })

summary = pd.DataFrame(rows)

# Rank models by average correct accuracy
model_order = (
    summary.groupby("Model")["Correct"]
    .mean()
    .sort_values(ascending=False)
    .index
    .tolist()
)

correct = (
    summary.pivot(index="Model", columns="Legal", values="Correct")
    .reindex(model_order)[LEGAL_ORDER]
)

compliant = (
    summary.pivot(index="Model", columns="Legal", values="Compliant")
    .reindex(model_order)[LEGAL_ORDER]
)

# Add average row
correct.loc["Average"] = correct.mean(axis=0)
compliant.loc["Average"] = compliant.mean(axis=0)

# Save underlying data
summary.to_csv(OUT_DIR / "legal_condition_correct_compliant_data.csv", index=False)

# ------------------------------------------------------------
# Plotting function
# ------------------------------------------------------------

def draw_heatmap(data, output_stem):
    fig, ax = plt.subplots(figsize=(7.2, 8.8), constrained_layout=True)

    im = ax.imshow(data.values, aspect="auto", vmin=VMIN, vmax=VMAX)

    ax.set_xticks(np.arange(len(data.columns)))
    ax.set_xticklabels(data.columns, rotation=25, ha="right")

    ax.set_yticks(np.arange(len(data.index)))
    ax.set_yticklabels(data.index)

    for i in range(data.shape[0]):
        for j in range(data.shape[1]):
            val = data.iloc[i, j]
            text_color = "white" if val < 55 else "black"

            ax.text(
                j,
                i,
                f"{val:.1f}%",
                ha="center",
                va="center",
                fontsize=8,
                fontweight="bold",
                color=text_color,
            )

    ax.tick_params(length=0)

    #cbar = fig.colorbar(im, ax=ax, shrink=0.8)
    #cbar.set_label("% of all responses")

    fig.savefig(OUT_DIR / f"{output_stem}.pdf", bbox_inches="tight")
    fig.savefig(OUT_DIR / f"{output_stem}.png", dpi=300, bbox_inches="tight")
    plt.close(fig)

# ------------------------------------------------------------
# Make the two separate heatmaps
# ------------------------------------------------------------

draw_heatmap(correct, "legal_conditions_correct_heatmap")
draw_heatmap(compliant, "legal_conditions_compliant_heatmap")

print(f"Saved charts to: {OUT_DIR.resolve()}")
print("Created:")
print(" - legal_conditions_correct_heatmap.pdf")
print(" - legal_conditions_compliant_heatmap.pdf")
print(" - legal_conditions_correct_heatmap.png")
print(" - legal_conditions_compliant_heatmap.png")