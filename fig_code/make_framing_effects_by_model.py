import re
from pathlib import Path

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

# ------------------------------------------------------------
# File paths
# ------------------------------------------------------------

INPUT_XLSX = "results_batch_run_6.xlsx"
OUT_DIR = Path("landlord_framing_charts")
OUT_DIR.mkdir(exist_ok=True)

# ------------------------------------------------------------
# Settings
# ------------------------------------------------------------

BASELINE = "A"              # no description
ILLEGAL_PREF = "C"          # illegal preference
NICE_PREF = "D"             # nice preference

UNDER_COLOR = "#d62728"     # red
OVER_COLOR = "#ff7f0e"      # default matplotlib orange

# ------------------------------------------------------------
# Load data
# ------------------------------------------------------------

input_path = Path(INPUT_XLSX)

if not input_path.exists():
    available = sorted(Path(".").glob("*.xlsx"))
    raise FileNotFoundError(
        f"Could not find {INPUT_XLSX} in the current folder.\n\n"
        "Excel files found here:\n"
        + "\n".join(str(p) for p in available)
    )

df = pd.read_excel(input_path)

if "Prompt ID" not in df.columns:
    raise ValueError("Missing required column: Prompt ID")

# ------------------------------------------------------------
# Extract landlord framing and correct answer from Prompt ID
# ------------------------------------------------------------

df["Landlord_code"] = (
    df["Prompt ID"]
    .astype(str)
    .str.extract(r"Landlord([A-F])")[0]
)

df["Problem_code"] = (
    df["Prompt ID"]
    .astype(str)
    .str.extract(r"_Problem[A-Z]_([123][A-Za-z]+)")[0]
    .str.lower()
)

missing_landlord = df[df["Landlord_code"].isna()]["Prompt ID"].head(10).tolist()
if missing_landlord:
    raise ValueError(
        "Could not identify landlord framing for some rows. Examples:\n"
        + "\n".join(map(str, missing_landlord))
    )

missing_problem = df[df["Problem_code"].isna()]["Prompt ID"].head(10).tolist()
if missing_problem:
    raise ValueError(
        "Could not identify problem code for some rows. Examples:\n"
        + "\n".join(map(str, missing_problem))
    )

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

model_names = []
for col in run_cols:
    match = re.match(r"(.+) R[1-5]$", col)
    if match:
        model_names.append(match.group(1))

model_names = list(dict.fromkeys(model_names))

# ------------------------------------------------------------
# Helper function: error rates for one model and one framing
# ------------------------------------------------------------

def get_rates(model, landlord_code):
    model_run_cols = [f"{model} R{i}" for i in range(1, 6)]

    missing = [c for c in model_run_cols if c not in df.columns]
    if missing:
        raise ValueError(f"Missing run columns for {model}: {missing}")

    sub = df[df["Landlord_code"] == landlord_code]

    total = 0
    under_count = 0
    over_count = 0
    correct_count = 0
    no_answer_count = 0

    for col in model_run_cols:
        responses = pd.to_numeric(sub[col], errors="coerce")
        correct_answer = sub["Correct_answer"]

        valid = responses.isin([1, 2, 3])

        correct = valid & (responses == correct_answer)
        under = valid & (responses < correct_answer)
        over = valid & (responses > correct_answer)
        no_answer = ~valid

        total += len(sub)
        correct_count += int(correct.sum())
        under_count += int(under.sum())
        over_count += int(over.sum())
        no_answer_count += int(no_answer.sum())

    return {
        "under": 100 * under_count / total,
        "over": 100 * over_count / total,
        "correct": 100 * correct_count / total,
        "no_answer": 100 * no_answer_count / total,
    }

# ------------------------------------------------------------
# Compute framing effects
# ------------------------------------------------------------

rows = []

for model in model_names:
    baseline = get_rates(model, BASELINE)
    illegal_pref = get_rates(model, ILLEGAL_PREF)
    nice_pref = get_rates(model, NICE_PREF)

    rows.append({
        "Model": model,
        "Illegal preference: change in under-recommendations":
            illegal_pref["under"] - baseline["under"],
        "Nice preference: change in over-recommendations":
            nice_pref["over"] - baseline["over"],
    })

summary = pd.DataFrame(rows)

# Sort by the larger of the two effects, so the most framing-sensitive models appear at top.
summary["Max absolute effect"] = summary[
    [
        "Illegal preference: change in under-recommendations",
        "Nice preference: change in over-recommendations",
    ]
].abs().max(axis=1)

summary = summary.sort_values("Max absolute effect", ascending=True)

summary.to_csv(
    OUT_DIR / "framing_effects_by_model_summary.csv",
    index=False,
)

print("\nFraming effects by model:")
print(
    summary[
        [
            "Model",
            "Illegal preference: change in under-recommendations",
            "Nice preference: change in over-recommendations",
        ]
    ].to_string(index=False)
)

# ------------------------------------------------------------
# Plot
# ------------------------------------------------------------

fig, ax = plt.subplots(figsize=(7.2, 5.2), constrained_layout=True)

y = np.arange(len(summary))

under_effect = summary["Illegal preference: change in under-recommendations"].values
over_effect = summary["Nice preference: change in over-recommendations"].values
models = summary["Model"].values

offset = 0.16

# Zero reference line
ax.axvline(0, linewidth=1, color="black", alpha=0.6)

# Dots
ax.scatter(
    under_effect,
    y + offset,
    s=48,
    color=UNDER_COLOR,
    marker="o",
)

ax.scatter(
    over_effect,
    y - offset,
    s=48,
    color=OVER_COLOR,
    marker="s",
)

# Value labels
for i, value in enumerate(under_effect):
    ha = "left" if value >= 0 else "right"
    dx = 0.7 if value >= 0 else -0.7
    ax.text(
        value + dx,
        i + offset,
        f"{value:+.1f}",
        va="center",
        ha=ha,
        fontsize=8,
        color=UNDER_COLOR,
    )

for i, value in enumerate(over_effect):
    ha = "left" if value >= 0 else "right"
    dx = 0.7 if value >= 0 else -0.7
    ax.text(
        value + dx,
        i - offset,
        f"{value:+.1f}",
        va="center",
        ha=ha,
        fontsize=8,
        color=OVER_COLOR,
    )

ax.set_yticks(y)
ax.set_yticklabels(models)

ax.set_xlabel("Percentage-point change relative to no description")

# Adjust if labels are clipped
x_min = min(under_effect.min(), over_effect.min(), 0) - 4
x_max = max(under_effect.max(), over_effect.max(), 0) + 6
ax.set_xlim(x_min, x_max)

ax.grid(axis="x", alpha=0.25)

ax.spines["top"].set_visible(False)
ax.spines["right"].set_visible(False)

# No title.
# No legend/key. Explain colours in the caption:
# red circles = illegal-preference effect on under-recommendations;
# orange squares = nice-preference effect on over-recommendations.

fig.savefig(
    OUT_DIR / "framing_effects_by_model_no_title.pdf",
    bbox_inches="tight",
)

fig.savefig(
    OUT_DIR / "framing_effects_by_model_no_title.png",
    dpi=300,
    bbox_inches="tight",
)

plt.close(fig)

print(f"\nSaved chart files to: {OUT_DIR.resolve()}")
print("Created:")
print(" - framing_effects_by_model_no_title.pdf")
print(" - framing_effects_by_model_no_title.png")
print(" - framing_effects_by_model_summary.csv")