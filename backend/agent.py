import json, os, re, time, asyncio
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import httpx

"""删除所有代理环境变量，防止请求通过代理出去"""
for _k in ["HTTP_PROXY", "HTTPS_PROXY", "http_proxy", "https_proxy"]:
    os.environ.pop(_k, None)

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "sk-b0faf54a9252450bbb7fb6e6fd56cf28")
DEEPSEEK_URL = "https://api.deepseek.com/chat/completions"
MODEL = "deepseek-chat"

# 共享的异步 HTTP 客户端（复用连接，避免每次重建）
_async_client: httpx.AsyncClient | None = None

def _get_async_client() -> httpx.AsyncClient:
    global _async_client
    if _async_client is None or _async_client.is_closed:
        _async_client = httpx.AsyncClient(timeout=90)
    return _async_client


async def _call_async(sys_msg: str, usr_msg: str, max_tokens: int = 2000) -> str:
    """异步调用 DeepSeek 流式 API，复用连接，重试最多3次"""
    h = {"Authorization": "Bearer " + DEEPSEEK_API_KEY, "Content-Type": "application/json"}
    p = {
        "model": MODEL, "temperature": 0.1, "max_tokens": max_tokens, "stream": True,
        "messages": [{"role": "system", "content": sys_msg}, {"role": "user", "content": usr_msg}],
    }
    client = _get_async_client()
    for i in range(3):
        try:
            content = ""
            async with client.stream("POST", DEEPSEEK_URL, headers=h, json=p) as r:
                r.raise_for_status()
                async for line in r.aiter_lines():
                    if not line or not line.startswith("data: "):
                        continue
                    data = line[6:]
                    if data.strip() == "[DONE]":
                        break
                    try:
                        chunk = json.loads(data)
                        content += chunk["choices"][0]["delta"].get("content", "")
                    except Exception:
                        pass
            return content
        except Exception:
            if i == 2:
                raise
            await asyncio.sleep(1)  # 比原来的 sleep(2) 更短
    return ""


def _call(sys_msg: str, usr_msg: str, max_tokens: int = 2000) -> str:
    """同步包装，供线程池内调用（线程池中不能直接 await）"""
    return asyncio.run(_call_async(sys_msg, usr_msg, max_tokens))


def _parse_json(text):
    text = text.strip()
    if "```" in text:
        parts = text.split("```")
        for part in parts:
            part = part.strip()
            if part.startswith("json"):
                part = part[4:].strip()
            if part.startswith("[") or part.startswith("{"):
                try:
                    return json.loads(part)
                except Exception:
                    pass
    try:
        return json.loads(text)
    except Exception:
        pass
    for ch in ["[", "{"]:
        idx = text.find(ch)
        if idx >= 0:
            try:
                return json.loads(text[idx:])
            except Exception:
                pass
    return []


async def _analyze_keywords_async(user_requirement, template_structure, source_samples=""):
    template_desc = json.dumps(template_structure, ensure_ascii=False)
    sys_msg = "You are a data analyst. Output JSON only, no explanation."
    usr_msg = (
        "Analyze the user requirement and template structure. "
        "For each table key, output filter conditions to select relevant rows from the data source.\n"
        "User requirement: " + user_requirement + "\n"
        "Template structure: " + template_desc + "\n"
        + (f"Data source sample (first few rows):\n{source_samples}\n" if source_samples else "")
        + "Output JSON like:\n"
        "{\n"
        '  "table_0": {"match_all": ["city_name", "date_str"], "match_any": []},\n'
        '  "Sheet1": {"match_all": [], "match_any": ["2020-07-", "2020-08-"]}\n'
        "}\n"
        "Rules:\n"
        "- match_all: row must contain ALL of these strings (AND logic)\n"
        "- match_any: row must contain ANY of these strings (OR logic, for date ranges)\n"
        "- For date ranges, look at the sample data to determine the actual date format\n"
        "- If no filter needed, use empty lists\n"
        "Output JSON only."
    )
    raw = await _call_async(sys_msg, usr_msg, max_tokens=800)
    result = _parse_json(raw)
    return result if isinstance(result, dict) else {}


def _filter_text(text, keywords):
    if not keywords:
        return text
    lines = text.split("\n")
    header_lines = [line for line in lines[:5] if "|" in line]
    matched_lines = [line for line in lines if any(kw in line for kw in keywords)]
    result = "\n".join(header_lines + matched_lines)
    print(f"  过滤: {len(text)} -> {len(result)} 字符 ({len(matched_lines)} 行匹配)")
    return result if matched_lines else text[:5000]


def _extract(chunk, headers, req):
    sys_msg = "You are a data extraction expert. Output JSON array only, no explanation."
    usr_msg = (
        "Extract all data rows matching these fields from the text below.\n"
        "Fields: " + json.dumps(headers, ensure_ascii=False) + "\n"
        "User requirement: " + req + "\n"
        "\n"
        "Normalization rules:\n"
        "1. First infer the real meaning of each field from the template field name, surrounding table context, and document context.\n"
        "2. For vague or inconsistent fields, output the value that best matches the business meaning.\n"
        "3. For region-related fields (地区, 省份, 所在地, 区域, 行政区划):\n"
        "   - Normalize to standard full administrative division name: 中国 + province full name.\n"
        "   - Examples: 广东 -> 中国广东省; 湖北 -> 中国湖北省.\n"
        "4. For auto-completion: if a field's value can be reliably inferred from other fields, fill it in.\n"
        "5. For time, units, numbers, and names: keep the original meaning unchanged, normalize format only.\n"
        "6. If a field cannot be determined reliably, output empty string.\n"
        "\n"
        "Text:\n" + chunk + "\n"
        "Output JSON array. Each element is an object with field names matching exactly. Output [] if no data."
    )
    raw = _call(sys_msg, usr_msg, max_tokens=8000)
    return _parse_json(raw)


async def _align_headers_async(source_headers, target_headers, sample_rows=None, user_requirement=""):
    sys_msg = "You are a data mapping expert. Output JSON only, no explanation."
    usr_msg = (
        "Map source spreadsheet headers to template headers.\n"
        "Source headers: " + json.dumps(source_headers, ensure_ascii=False) + "\n"
        "Template headers: " + json.dumps(target_headers, ensure_ascii=False) + "\n"
        + ("Sample rows from source sheet:\n" + json.dumps(sample_rows[:5], ensure_ascii=False) + "\n" if sample_rows else "")
        + ("User requirement: " + user_requirement + "\n" if user_requirement else "")
        + "Return a JSON object where keys are template headers and values are the matching source header names.\n"
        + "Rules:\n"
        + "- Prefer semantic matches over literal matches.\n"
        + "- If a template header is not represented in the source, use empty string.\n"
        + "- Do not invent source headers.\n"
        + "Output JSON only."
    )
    raw = await _call_async(sys_msg, usr_msg, max_tokens=1200)
    result = _parse_json(raw)
    if isinstance(result, dict):
        return {str(k): str(v) for k, v in result.items() if str(k).strip()}
    return {}


def _map_rows(src_rows, target_headers, header_map=None):
    result = []
    header_map = header_map or {}
    for src_row in src_rows:
        mapped = {}
        for th in target_headers:
            value = None
            mapped_src = header_map.get(th, "")
            if mapped_src:
                for sk, sv in src_row.items():
                    if sk == mapped_src or sk.replace(" ", "") == mapped_src.replace(" ", ""):
                        value = sv
                        break
            if value is None and th in src_row:
                value = src_row[th]
            if value is None:
                th_clean = th.replace(" ", "")
                for sk, sv in src_row.items():
                    if sk.replace(" ", "") == th_clean:
                        value = sv
                        break
            if value is not None:
                mapped[th] = value
        if any(v for v in mapped.values()):
            result.append(mapped)
    return result


async def extract_and_fill(source_texts, template_path, template_structure, user_requirement, source_paths=None):
    from extractor import extract_xlsx_rows

    has_large_file = any(
        (source_paths or {}).get(fname, "").lower().endswith((".xlsx", ".xls")) and (len(text) == 0 or len(text) > 50000)
        for fname, text in source_texts.items()
    )
    total_text_len = sum(len(t) for t in source_texts.values())

    # 预先缓存大 xlsx 的样本行，避免后续重复 load_workbook
    xlsx_sample_cache: dict[str, tuple] = {}  # fname -> (headers, sample_rows)

    keywords_map = {}
    if has_large_file or total_text_len > 15000:
        print("步骤1: 分析过滤关键词...")
        source_samples = ""
        for fname, text in source_texts.items():
            fpath = (source_paths or {}).get(fname, "")
            is_xlsx = fpath and Path(fpath).suffix.lower() in (".xlsx", ".xls")
            if is_xlsx and (len(text) == 0 or len(text) > 50000):
                try:
                    import openpyxl as _opx
                    _wb = _opx.load_workbook(fpath, data_only=True, read_only=True)
                    _ws = _wb.active
                    _lines = []
                    _sample_rows = []
                    _headers = None
                    for _i, _row in enumerate(_ws.iter_rows(values_only=True)):
                        row_data = [str(v) if v is not None else "" for v in _row]
                        if _i == 0:
                            _headers = row_data
                            _lines.append(" | ".join(row_data))
                        elif _i <= 3:
                            _lines.append(" | ".join(row_data))
                            _sample_rows.append(dict(zip(_headers or [], row_data)))
                        else:
                            break
                    _wb.close()
                    # 缓存样本，供后续 _align_headers_async 使用
                    xlsx_sample_cache[fname] = (_headers or [], _sample_rows)
                    source_samples += f"[{fname}]\n" + "\n".join(_lines) + "\n"
                except Exception:
                    pass
            elif text:
                source_samples += f"[{fname}]\n" + text[:300] + "\n"
        keywords_map = await _analyze_keywords_async(user_requirement, template_structure, source_samples)
        print("  关键词:", keywords_map)
    else:
        print("步骤1: 文本较小，跳过关键词分析，直接提取")

    # 并发处理每个 table key
    async def _process_key(key, data):
        headers = data.get("headers", [])
        if not headers:
            return key, []

        all_rows = []

        for fname, text in source_texts.items():
            fpath = (source_paths or {}).get(fname, "")
            is_xlsx = fpath and Path(fpath).suffix.lower() in (".xlsx", ".xls")

            if is_xlsx and (len(text) > 50000 or len(text) == 0):
                # 大 xlsx：结构化过滤 + 列映射
                filter_cond = keywords_map.get(key, {})
                match_all = filter_cond.get("match_all", []) if isinstance(filter_cond, dict) else (filter_cond or [])
                match_any = filter_cond.get("match_any", []) if isinstance(filter_cond, dict) else []

                loop = asyncio.get_event_loop()
                source_headers, src_rows = await loop.run_in_executor(
                    None, lambda: extract_xlsx_rows(fpath, match_all=match_all, match_any=match_any)
                )
                # 优先用缓存的样本行
                cached = xlsx_sample_cache.get(fname)
                sample_rows = cached[1] if cached else src_rows
                header_map = await _align_headers_async(source_headers, headers, sample_rows, user_requirement)
                print(f"  [{key}] {fname}: 表头映射 {header_map}")
                rows = _map_rows(src_rows, headers, header_map=header_map)
                print(f"  [{key}] {fname}: 语义映射 {len(rows)} 行")
                all_rows.extend(rows)
            else:
                # 文本文件或小 xlsx：关键词过滤 + LLM 提取
                filter_cond = keywords_map.get(key, {})
                kws = (filter_cond.get("match_all", []) + filter_cond.get("match_any", [])) if isinstance(filter_cond, dict) else (filter_cond or [])

                filtered = _filter_text(text, kws) if len(text) > 10000 else text

                # chunk 大小提升到 4000，减少 LLM 调用次数
                paragraphs = [p for p in filtered.split('\n') if p.strip()]
                chunks, cur = [], ""
                for p in paragraphs:
                    if len(cur) + len(p) > 2000 and cur:
                        chunks.append(cur)
                        cur = p
                    else:
                        cur += "\n" + p
                if cur:
                    chunks.append(cur)
                if not chunks:
                    chunks = [filtered]
                print(f"  [{key}] {fname}: {len(chunks)} 块")

                loop = asyncio.get_event_loop()
                with ThreadPoolExecutor(max_workers=min(len(chunks), 6)) as executor:
                    futures = [executor.submit(_extract, c, headers, user_requirement) for c in chunks]
                    chunk_results = await loop.run_in_executor(
                        None,
                        lambda fs=futures: [f.result() for f in as_completed(fs)]
                    )
                for rows in chunk_results:
                    if isinstance(rows, list):
                        all_rows.extend(rows)

        # 去重
        seen = set()
        unique = []
        for row in all_rows:
            k = str(row)
            if k not in seen:
                seen.add(k)
                unique.append(row)
        print(f"  [{key}] 提取到 {len(unique)} 行")
        return key, unique

    # 并发执行所有 table key 的提取
    tasks = [_process_key(key, data) for key, data in template_structure.items()]
    results = await asyncio.gather(*tasks)
    return dict(results)
