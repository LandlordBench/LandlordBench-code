import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ------------------------------------------------------------
# File paths
# ------------------------------------------------------------

INPUT_XLSX = "results_batch_run_6.xlsx"
OUT_DIR = Path("problem_charts")
OUT_DIR.mkdir(exist_ok=True)

# ------------------------------------------------------------
# Problem order
# ------------------------------------------------------------
# These are the exact labels used in the Prompt ID / filenames.

PROBLEM_ORDER = [
    "1fridge",
    "1lightbulb",
    "1stain",
    "2laminate",
    "2bathroomlock",
    "2badsmell",
    "3toilet",
    "3heater",
    "3suppliedfridge",
]

# ------------------------------------------------------------
# Load data
# ------------------------------------------------------------

df = pd.read_excel(INPUT_XLSX)

if "Prompt ID" not in df.columns:
    raise ValueError("Missing required column: Prompt ID")

# Extract problem code from Prompt ID.
# Example:
# LandlordA_LegalA_ProblemB_1lightbulb_ContextA
# becomes:
# 1lightbulb

df["Problem_code"] = (
    df["Prompt ID"]
    .astype(str)
    .str.extract(r"_Problem[A-Z]_([123][A-Za-z]+)")[0]
    .str.lower()
)

missing_problem_codes = sorted(set(df["Problem_code"].dropna()) - set(PROBLEM_ORDER))

if missing_problem_codes:
    raise ValueError(
        "Found problem codes that are not listed in PROBLEM_ORDER:\n"
        + "\n".join(missing_problem_codes)
        + "\n\nAdd them to PROBLEM_ORDER."
    )

# Extract correct answer from the first character of the problem code.
# For example:
# 1lightbulb -> 1
# 2badsmell -> 2
# 3heater -> 3

df["Correct_answer"] = df["Problem_code"].str.extract(r"^([123])")[0].astype(int)

# ------------------------------------------------------------
# Find model run columns
# ------------------------------------------------------------

run_cols = [
    col for col in df.columns
    if re.match(r".+ R[1-5]$", str(col))
]

if not run_cols:
    raise ValueError(
        "Could not find model run columns. Expected columns like 'GPT-5.4 R1'."
    )

# ------------------------------------------------------------
# Compute accuracy and compliance by problem
# ------------------------------------------------------------
# Accuracy:
#   response == correct answer
#
# Compliance:
#   response is correct or an over-recommendation.
#
# Since the options are ordered by legal strictness/cost:
#   1 = refuse
#   2 = slow/cheap repair
#   3 = fast/expensive repair
#
# A response is compliant if:
#   response >= correct answer
#
# No-answers, text answers, blanks, etc. are counted in the denominator
# but are neither correct nor compliant.

rows = []

for problem in PROBLEM_ORDER:
    sub = df[df["Problem_code"] == problem]

    if sub.empty:
        raise ValueError(f"No rows found for problem code: {problem}")

    total = 0
    correct_count = 0
    compliant_count = 0

    for col in run_cols:
        responses = pd.to_numeric(sub[col], errors="coerce")
        correct_answer = sub["Correct_answer"]

        valid = responses.isin([1, 2, 3])

        correct = valid & (responses == correct_answer)
        compliant = valid & (responses >= correct_answer)

        total += len(sub)
        correct_count += int(correct.sum())
        compliant_count += int(compliant.sum())

    rows.append({
        "Problem": problem,
        "Accuracy": 100 * correct_count / total,
        "Compliance": 100 * compliant_count / total,
        "N_responses": total,
        "Correct_count": correct_count,
        "Compliant_count": compliant_count,
    })

summary = pd.DataFrame(rows)

# Sort by compliance so the plot reads from low to high.
# Highest values will appear at the top of the chart.
summary = summary.sort_values("Compliance", ascending=True)

summary.to_csv(OUT_DIR / "problem_accuracy_compliance_summary.csv", index=False)

print("\nProblem summary:")
print(summary[["Problem", "Accuracy", "Compliance"]].to_string(index=False))

# ------------------------------------------------------------
# Draw dumbbell plot
# ------------------------------------------------------------

fig, ax = plt.subplots(figsize=(7.2, 4.8), constrained_layout=True)

y = np.arange(len(summary))

accuracy = summary["Accuracy"].values
compliance = summary["Compliance"].values
labels = summary["Problem"].values

# Connecting lines
ax.hlines(
    y=y,
    xmin=accuracy,
    xmax=compliance,
    linewidth=2,
    alpha=0.65,
)

# Points
ax.scatter(
    accuracy,
    y,
    s=55,
    label="Accuracy",
    zorder=3,
)

ax.scatter(
    compliance,
    y,
    s=55,
    marker="s",
    label="Compliance",
    zorder=3,
)

# Value labels
for i, (acc, comp) in enumerate(zip(accuracy, compliance)):
    if abs(comp - acc) < 1.0:
        ax.text(
            acc + 1.2,
            i,
            f"{acc:.1f}",
            va="center",
            ha="left",
            fontsize=9,
        )
    else:
        ax.text(
            acc - 1.2,
            i,
            f"{acc:.1f}",
            va="center",
            ha="right",
            fontsize=9,
        )
        ax.text(
            comp + 1.2,
            i,
            f"{comp:.1f}",
            va="center",
            ha="left",
            fontsize=9,
        )

ax.set_yticks(y)
ax.set_yticklabels(labels)

ax.set_xlabel("Percentage of all responses")
ax.set_xlim(20, 100)

ax.grid(axis="x", alpha=0.25)

ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

ax.legend(frameon=False, loc="lower right")

# No title or caption, because you will add those in Overleaf.
fig.savefig(OUT_DIR / "problem_accuracy_compliance_dumbbell.pdf", bbox_inches="tight")
fig.savefig(
    OUT_DIR / "problem_accuracy_compliance_dumbbell.png",
    dpi=300,
    bbox_inches="tight",
)

plt.close(fig)

print(f"\nSaved chart files to: {OUT_DIR.resolve()}")
print("Created:")
print(" - problem_accuracy_compliance_dumbbell.pdf")
print(" - problem_accuracy_compliance_dumbbell.png")
print(" - problem_accuracy_compliance_summary.csv")