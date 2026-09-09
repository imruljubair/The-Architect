<img width="1422" height="253" alt="image" src="https://github.com/user-attachments/assets/88e2951a-f5a9-4711-8c56-d8dd0ada6d9a" />

# The Architect
Interactive Visualization of Deep Learning Mathematics Directly in Microsoft Excel.

- Technical report: https://arxiv.org/abs/2608.13572
- Youtube Tutorial: https://youtu.be/HzAudh_PK8I?si=mjjd1A8EpEM2QA6G

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

<img width="1324" height="736" alt="image" src="https://github.com/user-attachments/assets/c81ae6d4-d1de-4969-86d6-645005b82c64" />

<img width="1324" height="736" alt="image" src="https://github.com/user-attachments/assets/cb9eb5ba-f775-48cf-9ab1-7576f3000f83" />

