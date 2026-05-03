# LandlordBench Code

## Overview

This repository contains the code used to generate prompts, construct legal resources, and run model evaluations for **LandlordBench**, a benchmark for evaluating the legal compliance of large language models (LLMs) on tenancy repair decisions.

The associated dataset is available separately.

---

## Repository Structure

* `generate_prompts.py`
  Generates the full set of prompts by combining landlord framing, legal resources, repair scenarios, and tenant pressure.

* `cases/`
  Contains the legal case materials used in the “schedule + cases” condition.

* `batch_run.py`
  Executes model calls over all prompts, records outputs, and logs metadata.

---

## Functionality

The code implements the following pipeline:

1. **Prompt generation**
   Creates structured prompts using a fixed template and systematically varied conditions.

2. **Legal resource integration**
   Loads statutory text and case law into prompts depending on the experimental condition.

3. **Batch execution**
   Sends prompts to multiple models across repeated runs and records outputs.

4. **Logging and storage**
   Saves raw model outputs, parsed responses, and metadata for evaluation.

---

## Requirements

* Python 3.10+
* Standard libraries (`os`, `csv`, `json`, etc.)
* API client for model access (e.g., OpenAI-compatible client via OpenRouter)

---

## Usage

### 1. Generate prompts

```bash
python generate_prompts.py
```

This creates the full set of prompt `.txt` files.

---

### 2. Run model evaluations

```bash
python batch_run.py
```

This:

* sends prompts to models
* records responses
* stores results in CSV / Excel formats

---

## Data

The dataset used in this project is hosted separately and includes:

* `landlordbench.csv` (prompts + labels)
* `full_results.csv` (model outputs and evaluation data)

---

## Reproducibility

The repository contains all code necessary to:

* regenerate prompts
* run model evaluations
* reproduce the experimental pipeline

---

## License

This code is released under the Creative Commons Attribution 4.0 International License (CC BY 4.0).

---

## Anonymity Notice

This repository is provided in anonymised form for peer review.
