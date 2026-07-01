# ROST Productivity Shiny App

This bundle contains the files needed to run the uploadable, multi-year
ROST productivity dashboard. It does not contain project source data.

## Requirements

- Python 3.11 or newer
- Internet access during the initial package installation

## Setup

Open a terminal in this folder and run:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

On Windows PowerShell, activate the environment with:

```powershell
.\.venv\Scripts\Activate.ps1
```

## Launch

From this folder, run:

```bash
shiny run --reload --launch-browser shiny/app.py
```

If the browser does not open automatically, visit:

```text
http://127.0.0.1:8000
```

Stop the app with `Ctrl+C` in the terminal.

## Using the app

1. Upload a productivity `.xlsx` workbook.
2. Optionally upload a resight `.xlsx` workbook.
3. Select the relevant worksheet in each workbook.
4. Review or change the column mappings.
5. Map a species column, or enter one species label for the whole workbook.
6. Confirm the detected analysis years.
7. Select **Generate analysis**.
8. Use the year and species selectors to explore a single season.
9. Use **Compare** for same-species cross-year or same-year cross-species charts.
10. Use **Downloads** to export a selected PNG/CSV or a ZIP of all results.

Uploads are temporary and remain isolated to the browser session. The app does
not save uploaded workbooks or generated analysis tables.

## Included files

```text
manager_shiny_bundle/
├── README.md
├── requirements.txt
├── scripts/
│   └── clean_productivity.py
└── shiny/
    ├── app.py
    ├── data.py
    └── www/
        └── styles.css
```
