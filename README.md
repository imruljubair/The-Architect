<img width="1422" height="253" alt="image" src="https://github.com/user-attachments/assets/88e2951a-f5a9-4711-8c56-d8dd0ada6d9a" />

# The Architect

Interactive forward-pass deep-learning mathematics in Microsoft Excel, preserving the original workbook layout
and all three original buttons: **🎲 Randomize**, **↺ Clear**, and **▶ Run**.
The backpropagation checkbox is disabled, and backward-pass, parameter-update,
training-epoch, and dashboard generation code is removed. Forward loss calculations
and PyTorch forward code remain.

Open **theArchitect_finalV2.xlsx** in this directory and use the original buttons.
The Run button targets the dedicated **The Architect - Forward Pass** Office Script.
Randomize and Clear retain their original script links and behavior.

The source is provided in `src/`; the original `.osts` filenames remain at the
repository root. To install under another Microsoft 365 account, save each script
in Automate and associate the respective worksheet button with it. Office Scripts
are stored in OneDrive and referenced by workbook buttons.

For development, run `npm ci` and `npm run check`. `npm run build` copies the
editable TypeScript sources into the distributed `.osts` files.

## Contents

- [Excel workbook](theArchitect_finalV2.xlsx): the forward-only workbook.
- `src/`: editable TypeScript for Run, Randomize, and Clear.
- Root `.osts` files: distributable Office Scripts.
- `tests/`: forward-generation and saved-workbook checks.

Backpropagation, parameter updates, and training diagnostics are available in the
separate **Hack Propagation** project.

## References

- [Technical report](https://arxiv.org/abs/2608.13572)
- [Video tutorial](https://youtu.be/HzAudh_PK8I)

The original report and tutorial describe the combined forward and training artifact.
This repository provides the forward-only edition.
