## Excel to PPT Agent

This project analyzes Excel data, generates insights and editable charts (embedded Excel data so charts are editable in PowerPoint), and builds a PowerPoint you can download via an API or CLI.

### Features
- Upload an Excel file; we infer date, numeric and categorical columns
- Generate key insights (totals, trends, top categories)
- Create editable charts in PowerPoint (line, bar, clustered bar)
- Download the generated `.pptx`

### Setup
1. Python 3.10+ is required (works on 3.13 without building pandas)
2. Create venv and install requirements:
   - `python3 -m venv excel_to_ppt_agent/.venv`
   - `source excel_to_ppt_agent/.venv/bin/activate`
   - `pip install -r excel_to_ppt_agent/requirements.txt`

### Run API
- `uvicorn app.api:app --reload --host 0.0.0.0 --port 8000`
- Then POST an Excel file to `/upload` and use the returned `download` path

Example curl:
```bash
curl -F "file=@excel_to_ppt_agent/sample_data/sales.xlsx" http://localhost:8000/upload
```

### CLI Usage
- `python excel_to_ppt_agent/cli.py /path/to/file.xlsx -o /path/to/report.pptx`

Outputs are saved to `excel_to_ppt_agent/outputs`.

### Notes
- Charts are editable in PowerPoint because they are created with `python-pptx` `add_chart` and embedded series data.
- We avoid heavy dependencies and read Excel via `openpyxl` for broad compatibility.

