"""
FastAPI 主入口 - 提供异步API接口
"""
import asyncio
import os
import time
import uuid
import shutil
from pathlib import Path
from typing import Optional
from functools import partial

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles

from extractor import extract_file, get_xlsx_structure, get_docx_structure
from agent import extract_and_fill
from filler import fill_xlsx, fill_docx

app = FastAPI(title="智能文档填表系统", version="1.0.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# 存储上传的数据源文档内容 {filename: text}
SOURCE_DOCS: dict[str, str] = {}
# 存储上传的数据源文件路径 {filename: path}
SOURCE_PATHS: dict[str, str] = {}

UPLOAD_DIR = Path("uploads")
OUTPUT_DIR = Path("../outputs")
FRONTEND_DIR = Path("../frontend")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

"""
当访问 http://host:8000/ 时，返回 ../frontend/index.html 文件
"""
@app.get("/")
def root():
    return FileResponse(str(FRONTEND_DIR / "index.html"))


app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")

#上传数据源文件
@app.post("/api/upload-sources")
async def upload_sources(files: list[UploadFile] = File(...)):
    """批量上传数据源文档（docx/xlsx/md/txt），提取文本存入内存"""
    results = []
    for file in files:
        suffix = Path(file.filename).suffix.lower()
        if suffix not in (".docx", ".xlsx", ".xls", ".md", ".txt"):
            results.append({"filename": file.filename, "status": "跳过", "reason": "不支持的格式"})
            continue

        save_path = UPLOAD_DIR / file.filename
        with open(save_path, "wb") as f:
            f.write(await file.read())

        try:
            # 大 xlsx 文件跳过全文提取，只存路径，后续按需过滤
            import os
            file_size = os.path.getsize(str(save_path))
            if suffix in (".xlsx", ".xls") and file_size > 1 * 1024 * 1024:  # 1MB
                SOURCE_DOCS[file.filename] = ""  # 占位
                SOURCE_PATHS[file.filename] = str(save_path)
                results.append({
                    "filename": file.filename,
                    "status": "成功",
                    "chars": 0,
                    "note": "大文件，将在填表时按需过滤"
                })
            else:
                loop = asyncio.get_event_loop()
                text = await loop.run_in_executor(None, extract_file, str(save_path))
                SOURCE_DOCS[file.filename] = text
                SOURCE_PATHS[file.filename] = str(save_path)
                results.append({
                    "filename": file.filename,
                    "status": "成功",
                    "chars": len(text)
                })
        except Exception as e:
            results.append({"filename": file.filename, "status": "失败", "reason": str(e)})

    return {
        "uploaded": len([r for r in results if r["status"] == "成功"]),
        "total_sources": len(SOURCE_DOCS),
        "details": results
    }


"""
查看已上传数据源
返回当前内存中存储的所有数据源文件名及其字符数（0 表示大文件未提取全文）
"""
@app.get("/api/sources")
def list_sources():
    """查看已上传的数据源文档列表"""
    return {
        "count": len(SOURCE_DOCS),
        "files": [
            {"filename": k, "chars": len(v)}
            for k, v in SOURCE_DOCS.items()
        ]
    }

"""
清空数据源
"""
@app.delete("/api/sources")
def clear_sources():
    """清空数据源"""
    SOURCE_DOCS.clear()
    SOURCE_PATHS.clear()
    return {"status": "已清空"}


"""
填表接口
"""
@app.post("/api/fill-template")
async def fill_template(
    template: UploadFile = File(...),
    requirement: Optional[str] = Form(default="智能填表，根据数据源内容填写模板中的所有字段"),
):
    """
    上传模板文件，系统自动从已上传的数据源中提取数据并填写模板。
    返回填写完成的文件。
    """
    if not SOURCE_DOCS:
        raise HTTPException(status_code=400, detail="请先上传数据源文档")

    start_time = time.time()

    # 保存模板文件
    template_filename = template.filename
    template_path = UPLOAD_DIR / f"template_{uuid.uuid4().hex[:8]}_{template_filename}"
    with open(template_path, "wb") as f:
        f.write(await template.read())

    ext = Path(template_filename).suffix.lower()
    if ext not in (".xlsx", ".xls", ".docx"):
        raise HTTPException(status_code=400, detail="模板文件必须是 .xlsx 或 .docx 格式")

    try:
        # 获取模板结构（阻塞 IO，放到线程池）
        loop = asyncio.get_event_loop()
        if ext in (".xlsx", ".xls"):
            template_structure = await loop.run_in_executor(None, get_xlsx_structure, str(template_path))
        else:
            raw = await loop.run_in_executor(None, get_docx_structure, str(template_path))
            # 转换为 {table_0: {headers, description}, ...} 格式
            paragraphs = raw.get("paragraphs", [])
            tables = raw.get("tables", [])
            template_structure = {}
            for t in tables:
                idx = t["index"]
                # 段落顺序：第0段是标题，第1段对应表0，第2段对应表1...
                desc = paragraphs[idx + 1] if idx + 1 < len(paragraphs) else ""
                template_structure[f"table_{idx}"] = {
                    "headers": t["headers"],
                    "description": desc,
                    "row_count": t.get("row_count", 0),
                }

        # 调用agent提取并生成填表数据（异步）
        fill_data = await extract_and_fill(
            source_texts=SOURCE_DOCS,
            source_paths=SOURCE_PATHS,
            template_path=str(template_path),
            template_structure=template_structure,
            user_requirement=requirement,
        )

        # 写入模板
        output_filename = f"filled_{uuid.uuid4().hex[:8]}_{template_filename}"
        output_path = OUTPUT_DIR / output_filename

        if ext in (".xlsx", ".xls"):
            await loop.run_in_executor(None, fill_xlsx, str(template_path), str(output_path), fill_data)
        else:
            await loop.run_in_executor(None, fill_docx, str(template_path), str(output_path), fill_data)

        elapsed = time.time() - start_time

        return {
            "status": "成功",
            "output_file": output_filename,
            "elapsed_seconds": round(elapsed, 2),
            "download_url": f"/api/download/{output_filename}",
            "fill_data_preview": {
                k: v[:2] if isinstance(v, list) else v
                for k, v in fill_data.items()
            }
        }

    except Exception as e:
        raise HTTPException(status_code=500, detail=f"填表失败: {str(e)}")
    finally:
        # 清理临时模板文件
        if template_path.exists():
            template_path.unlink()


"""
下载填写后文件接口
"""
@app.get("/api/download/{filename}")
def download_file(filename: str):
    """下载填写完成的文件"""
    # 防止路径穿越
    safe_name = Path(filename).name
    file_path = OUTPUT_DIR / safe_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="文件不存在")
    return FileResponse(
        path=str(file_path),
        filename=safe_name,
        media_type="application/octet-stream"
    )


@app.get("/api/outputs")
def list_outputs():
    """列出所有已生成的输出文件"""
    files = []
    for f in OUTPUT_DIR.iterdir():
        if f.is_file():
            files.append({"filename": f.name, "size_kb": round(f.stat().st_size / 1024, 1)})
    return {"files": files}


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
