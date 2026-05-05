# LandlordBench Code

This repository contains code and supporting materials for LandlordBench, a benchmark for evaluating whether large language models can select the cheapest legally permissible response to residential tenancy repair requests.

The benchmark concerns repair obligations under Australian Capital Territory tenancy law.

## Repository structure

```text
LandlordBench-code/
├── Batch_run_6_3.py
├── generate_prompts.py
├── Licence.txt
├── README.txt
├── clean_code/
│   └── build_human_dataset_annotated.py
├── fig_code/
│   ├── results_batch_run_6.xlsx
│   ├── human_results_batch_pooled_B.xlsx
│   ├── make_overall_response_stacked_bars.py
│   ├── make_problem_accuracy_compliance_dumbbell.py
│   ├── make_landlord_framing_lineplot.py
│   ├── make_framing_effects_by_model.py
│   ├── Make_legal_condition_charts.py
│   ├── make_human_ai_by_problem.py
│   └── make_human_model_by_landlord_framing.py
└── materials/
    ├── Residential Tenancies Act 1997.docx
    ├── case files used as legal materials
    └── generate_prompts.py
