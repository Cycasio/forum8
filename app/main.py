import json
import shutil
import subprocess
from datetime import datetime
from pathlib import Path
from typing import Any, Dict

from docxtpl import DocxTemplate
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse
from jinja2 import Environment, StrictUndefined
from starlette.background import BackgroundTask

BASE_DIR = Path(__file__).resolve().parent.parent
UPLOAD_DIR = BASE_DIR / "data" / "uploads"
OUTPUT_DIR = BASE_DIR / "data" / "outputs"
TMP_DIR = BASE_DIR / "data" / "tmp"
TEMPLATE_NAME = "template.docx"

for directory in (UPLOAD_DIR, OUTPUT_DIR, TMP_DIR):
    directory.mkdir(parents=True, exist_ok=True)

app = FastAPI(title="Fixed Layout Document Generator")


def render_homepage() -> str:
    template_uploaded = (UPLOAD_DIR / TEMPLATE_NAME).exists()
    return f"""
    <html>
        <head>
            <title>Fixed Layout Word Generator</title>
            <style>
                body {{ font-family: Arial, sans-serif; max-width: 960px; margin: 40px auto; }}
                form {{ margin-bottom: 30px; padding: 20px; border: 1px solid #ddd; border-radius: 8px; }}
                label {{ display: block; margin-top: 10px; }}
                textarea {{ width: 100%; height: 220px; font-family: monospace; }}
                .status {{ padding: 10px; background: #f3f3f3; border-radius: 6px; margin-bottom: 16px; }}
                .note {{ color: #666; font-size: 14px; }}
            </style>
        </head>
        <body>
            <h1>固定版型 Word 文件產生器</h1>
            <div class="status">目前母版狀態：{'已上傳' if template_uploaded else '尚未上傳'}</div>

            <h2>上傳母版 template.docx</h2>
            <form action="/upload-template" method="post" enctype="multipart/form-data">
                <label>選擇 template.docx：</label>
                <input type="file" name="file" accept="application/vnd.openxmlformats-officedocument.wordprocessingml.document" required />
                <button type="submit">上傳並覆蓋</button>
            </form>

            <h2>產生文件</h2>
            <form action="/generate" method="post">
                <label>JSON 內容（key 對應 {{key}}）：</label>
                <textarea name="json_data" required>{{{{
    "dept_name": "研發部",
    "requester": "王小明"
}}}}</textarea>
                <label><input type="checkbox" name="generate_pdf" /> 同時輸出 PDF</label>
                <button type="submit">產生並下載</button>
                <p class="note">若 JSON 缺少母版中的占位符會回報錯誤。</p>
            </form>
        </body>
    </html>
    """


@app.get("/", response_class=HTMLResponse)
async def home() -> str:
    return render_homepage()


@app.post("/upload-template")
async def upload_template(file: UploadFile = File(...)) -> JSONResponse:
    if not file.filename.lower().endswith(".docx"):
        raise HTTPException(status_code=400, detail="僅接受 .docx 母版檔案")

    destination = UPLOAD_DIR / TEMPLATE_NAME
    with destination.open("wb") as buffer:
        shutil.copyfileobj(file.file, buffer)

    return JSONResponse({"message": "母版已更新", "path": str(destination)})


def _load_context_from_request(content_type: str, request_data: Any) -> tuple[Dict[str, Any], bool]:
    if "application/json" in content_type:
        if not isinstance(request_data, dict):
            raise HTTPException(status_code=400, detail="JSON 請求格式錯誤")
        context = request_data.get("data")
        generate_pdf = bool(request_data.get("generate_pdf", False))
        if not isinstance(context, dict):
            raise HTTPException(status_code=400, detail="data 欄位必須是 JSON 物件")
        return context, generate_pdf

    if not hasattr(request_data, "get"):
        raise HTTPException(status_code=400, detail="表單資料解析失敗")

    json_text = request_data.get("json_data")
    if not json_text:
        raise HTTPException(status_code=400, detail="缺少 json_data")
    try:
        context = json.loads(json_text)
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=400, detail=f"JSON 解析錯誤: {exc}") from exc

    generate_pdf = request_data.get("generate_pdf") in {"on", True, "true"}
    if not isinstance(context, dict):
        raise HTTPException(status_code=400, detail="JSON 根層級必須是物件")
    return context, generate_pdf


def _render_docx(context: Dict[str, Any]) -> Path:
    template_path = UPLOAD_DIR / TEMPLATE_NAME
    if not template_path.exists():
        raise HTTPException(status_code=400, detail="尚未上傳母版 template.docx")

    try:
        document = DocxTemplate(template_path)
    except Exception as exc:  # pragma: no cover - defensive
        raise HTTPException(status_code=500, detail=f"無法載入母版: {exc}") from exc

    jinja_env = Environment(autoescape=False, undefined=StrictUndefined)
    try:
        document.render(context, jinja_env=jinja_env)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"產生內容失敗: {exc}") from exc

    timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
    output_name = f"output_{timestamp}.docx"
    output_path = OUTPUT_DIR / output_name
    document.save(output_path)
    return output_path


def _convert_to_pdf(docx_path: Path) -> Path:
    pdf_name = docx_path.with_suffix(".pdf").name
    pdf_path = OUTPUT_DIR / pdf_name

    command = [
        "soffice",
        "--headless",
        "--convert-to",
        "pdf",
        str(docx_path),
        "--outdir",
        str(OUTPUT_DIR),
    ]

    result = subprocess.run(command, capture_output=True, text=True)
    if result.returncode != 0:
        raise HTTPException(
            status_code=500,
            detail=f"PDF 轉換失敗: {result.stderr or result.stdout}",
        )

    if not pdf_path.exists():
        raise HTTPException(status_code=500, detail="PDF 檔案未生成")

    return pdf_path


@app.post("/generate")
async def generate(request: Request):
    content_type = request.headers.get("content-type", "")

    if "application/json" in content_type:
        payload = await request.json()
        context, generate_pdf = _load_context_from_request(content_type, payload)
    else:
        form_data = await request.form()
        context, generate_pdf = _load_context_from_request(content_type, form_data)

    docx_path = _render_docx(context)

    if not generate_pdf:
        return FileResponse(
            path=docx_path,
            media_type=(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            ),
            filename=docx_path.name,
        )

    pdf_path = _convert_to_pdf(docx_path)

    # package into zip so both docx and pdf are downloadable together
    timestamp = docx_path.stem.replace("output_", "")
    zip_path = TMP_DIR / f"output_{timestamp}.zip"
    with zip_path.open("wb") as zip_file:
        import zipfile

        with zipfile.ZipFile(zip_file, "w") as archive:
            archive.write(docx_path, arcname=docx_path.name)
            archive.write(pdf_path, arcname=pdf_path.name)

    cleanup_task = BackgroundTask(lambda: zip_path.unlink(missing_ok=True))
    return FileResponse(
        path=zip_path,
        media_type="application/zip",
        filename=zip_path.name,
        background=cleanup_task,
    )


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=8000, reload=True)
