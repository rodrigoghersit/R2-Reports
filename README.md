# R2 Reports

This repository contains a Python script that builds a LaTeX report from an Excel workbook.

## Folder structure

```
.
├── Report Generation.py
├── data/                 # place Excel input files here
└── Figures/
    ├── Overlays/         # PNG files referenced by the report
    └── Plots/            # Plot images referenced by the report
```

Running the script will create a `Matter` directory containing the generated LaTeX sections and PDF summaries. The main `*.tex` file is written to the repository root (or any path you choose).

## Usage

Ensure the required Python packages are installed:

```bash
pip install pandas
```

Then execute the script:

```bash
python "Report Generation.py"
```

The example at the bottom of the script assumes an Excel file called `Kingaroy SF HP2 Testing.xlsx` is located in the `data` folder.
