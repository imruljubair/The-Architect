# The-Architect
Interactive Visualization of Deep Learning Mathematics Directly in Microsoft Excel.

Technical report: https://arxiv.org/abs/2608.13572

Youtube Tutorial: https://youtu.be/HzAudh_PK8I?si=mjjd1A8EpEM2QA6G

## Contents

- `Code/theArchitect_v2.xlsx`: Excel workbook containing the Architect interface, generated blueprint views, and example data.
- `Code/architect_Mar24_2026_improve.osts`: Office Script source used to generate and update workbook blueprints.

## What This Artifact Demonstrates

The artifact supports the workflow described in the paper:

- specifying a neural network in the Architect tab
- generating forward and backward spreadsheet blueprints
- materializing multi-epoch training views
- inspecting and editing formulas directly inside Excel after generation

## Requirements

- Microsoft Excel with Office Scripts support
- A recent Excel version with dynamic array formulas enabled

## How To Use

1. Open `Code/theArchitect_v2.xlsx` in Microsoft Excel.
2. Open the Office Script editor.
3. Import or paste the script from `Code/architect_Mar24_2026_improve.osts`.
4. In the workbook, go to the Architect tab and edit the network specification.
5. Run the script to generate the blueprint and, when enabled, epoch sheets and dashboard views.
6. After generation, you can continue investigating by editing workbook cells and formulas directly.

## Notes For Review

- This repository is intentionally anonymized.
- The workbook may contain example configurations and generated sheets used for demonstration.
