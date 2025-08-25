from fastapi import FastAPI, UploadFile, File
from fastapi.responses import FileResponse, JSONResponse
import tempfile, os
from pathlib import Path
from urllib.parse import quote

from app.insights import analyze_excel
from app.ppt_builder import build_ppt_from_analysis

OUTPUTS_DIR = Path("/workspace/excel_to_ppt_agent/outputs")
OUTPUTS_DIR.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Excel to PPT Agent")


@app.post("/upload")
async def upload_excel(file: UploadFile = File(...)):
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        content = await file.read()
        tmp.write(content)
        tmp_path = tmp.name

    try:
        analysis = analyze_excel(tmp_path)
        ppt_path = build_ppt_from_analysis(analysis)
        ppt_name = Path(ppt_path).name
        return JSONResponse({"download": f"/download/{quote(ppt_name)}"})
    finally:
        os.unlink(tmp_path)


@app.get("/download/{filename}")
async def download(filename: str):
    safe_path = OUTPUTS_DIR / filename
    if not safe_path.exists():
        return JSONResponse({"error": "File not found"}, status_code=404)
    return FileResponse(str(safe_path), filename=safe_path.name)
