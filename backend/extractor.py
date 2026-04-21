"""
文档内容提取模块 - 支持 .docx / .xlsx / .md / .txt,将整个文件内容提取为纯文本(用于后续 LLM 处理)
对于 Excel 文件，提供按关键词过滤行并返回结构化数据（表头 + 行字典）的能力
提供获取模板结构信息的功能（例如 Excel 的表头、Word 的段落和表格表头），用于后续填充
"""
import re
from pathlib import Path
"""读写xslx文件"""
import openpyxl
"""读写word文件"""
from docx import Document

"""
从word文件中提取所有文本,包括段落和表格内容
返回一个字符串，段落和表格行之间用换行分隔，表格内单元格用 | 连接
"""
def extract_text_from_docx(file_path: str) -> str:
    doc = Document(file_path)
    # 存储提取的文本片段
    parts = []
    # 遍历所有段落,把清理干净的文本，添加到 parts
    for para in doc.paragraphs:
        if para.text.strip():
            parts.append(para.text.strip())
    # 遍历所有表格
    for table in doc.tables:
        for row in table.rows:
            # 提取一行中每个单元格的文本，并去除首尾空格
            row_data = [cell.text.strip() for cell in row.cells]
            if any(row_data):
                parts.append(" | ".join(row_data))
    # 将所有片段用换行符连接
    return "\n".join(parts)

"""
从 Excel 文件中提取所有工作表的内容为纯文本
每个工作表以 [Sheet: 名称] 开头，每行数据用 | 连接
"""
def extract_text_from_xlsx(file_path: str) -> str:
    # data_only=True 表示获取公式的计算值而非公式本身
    wb = openpyxl.load_workbook(file_path, data_only=True)
    parts = []
    # 遍历所有工作表
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        parts.append(f"[Sheet: {sheet_name}]")
        # 逐行读取单元格的值，转换成字符串
        for row in ws.iter_rows(values_only=True):
            row_data = [str(v) if v is not None else "" for v in row]
            if any(v.strip() for v in row_data):
                parts.append(" | ".join(row_data))
    return "\n".join(parts)

"""读取 Markdown 文件，返回原始文本内容"""
def extract_text_from_md(file_path: str) -> str:
    # 以 UTF-8 编码打开，忽略无法解码的字符
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

"""读取 txt 文件，返回原始文本内容"""
def extract_text_from_txt(file_path: str) -> str:
    with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
        return f.read()

"""从 Excel 文件中按关键词条件过滤行，返回 (headers, rows)"""
def extract_xlsx_rows(file_path: str, match_all: list = None, match_any: list = None) -> tuple:
    """按条件过滤xlsx，返回 (headers, rows) 结构化数据。
    match_all: 行必须包含所有关键词（AND）
    match_any: 行包含任意关键词即可（OR）
    两者都有时：match_all AND match_any 都满足"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    all_headers = []#and逻辑的表头
    all_rows = []#存储and逻辑匹配行
    #遍历所有sheet
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = None
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            row_data = [str(v) if v is not None else "" for v in row]
            #第一行作为表头（最终返回的 all_headers 只是第一个 sheet 的表头）
            if i == 0:
                headers = row_data
                if not all_headers:
                    all_headers = headers
                continue
            # 将整行所有单元格内容拼接成一个字符串，空格分隔
            row_str = " ".join(row_data)
            ok = True
            #如果 match_all 存在，要求所有关键词都出现在 row_str 中
            if match_all:
                ok = ok and all(kw in row_str for kw in match_all)
            #如果 match_any 存在，要求至少一个关键词出现
            #两者同时存在时，and 连接，即行必须同时满足两组条件
            if match_any:
                ok = ok and any(kw in row_str for kw in match_any)
            #存储匹配行到all_rows
            if ok:
                all_rows.append(dict(zip(headers, row_data)))
    print(f"  xlsx结构化过滤: 匹配{len(all_rows)}行")
    #返回第一个 sheet 的表头，以及所有匹配行的字典列表
    return all_headers, all_rows


"""根据文件扩展名自动选择对应的提取函数"""
def extract_file(file_path: str) -> str:
    path = Path(file_path)
    ext = path.suffix.lower()
    if ext == ".docx":
        return extract_text_from_docx(file_path)
    elif ext in (".xlsx", ".xls"):
        return extract_text_from_xlsx(file_path)
    elif ext == ".md":
        return extract_text_from_md(file_path)
    elif ext == ".txt":
        return extract_text_from_txt(file_path)
    else:
        raise ValueError(f"不支持的文件类型: {ext}")


"""获取xlsx和word模板的结构信息"""

def get_xlsx_structure(file_path: str) -> dict:
    """获取xlsx模板的结构信息（sheet名、表头、行数）"""
    wb = openpyxl.load_workbook(file_path, data_only=True)
    structure = {}
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        headers = []
        rows = []
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            row_data = [str(v) if v is not None else "" for v in row]
            if i == 0:
                headers = row_data
            else:
                if any(v.strip() for v in row_data):
                    rows.append(row_data)
        structure[sheet_name] = {"headers": headers, "existing_rows": rows}
    return structure

"""获取 Word 模板的结构信息,{"paragraphs":[段落1,段落2,...],"tables":[{"index":0,"headers":[表头1,表头2,...],"row_counts":5}]}"""
def get_docx_structure(file_path: str) -> dict:
    """获取docx模板的结构信息（段落描述、表格表头）"""
    doc = Document(file_path)
    paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
    tables = []
    for i, table in enumerate(doc.tables):
        if table.rows:
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            tables.append({"index": i, "headers": headers, "row_count": len(table.rows)})
    return {"paragraphs": paragraphs, "tables": tables}
