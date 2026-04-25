"""
pytest 配置文件：服务器启动/关闭 fixture、全局路径常量
"""
import os
import subprocess
import sys
import time
import threading
from pathlib import Path

import httpx
import pytest

# ── 路径常量 ──────────────────────────────────────────────────────────────────
ROOT = Path(__file__).parent.parent
BACKEND_DIR = ROOT / "backend"
TEST_DATA = ROOT / "测试集"
SCENARIO_DIR = TEST_DATA / "包含模板文件"

SCENARIO_COVID = SCENARIO_DIR / "COVID-19数据集"
SCENARIO_SHANDONG = SCENARIO_DIR / "2025山东省环境空气质量监测数据信息"
SCENARIO_CITY = SCENARIO_DIR / "2025年中国城市经济百强全景报告"

BASE_URL = "http://127.0.0.1:18000"

# ── 服务器 fixture ────────────────────────────────────────────────────────────

@pytest.fixture(scope="session")
def server():
    """在测试会话期间启动 FastAPI 服务器，结束后关闭。"""
    env = os.environ.copy()
    env["PYTHONPATH"] = str(BACKEND_DIR)

    proc = subprocess.Popen(
        [
            sys.executable, "-m", "uvicorn",
            "main:app",
            "--host", "127.0.0.1",
            "--port", "18000",
            "--log-level", "warning",
        ],
        cwd=str(BACKEND_DIR),
        env=env,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    # 等待服务就绪（最多 30 秒）
    deadline = time.time() + 30
    while time.time() < deadline:
        try:
            r = httpx.get(f"{BASE_URL}/", timeout=2)
            if r.status_code < 500:
                break
        except Exception:
            time.sleep(0.5)
    else:
        proc.terminate()
        raise RuntimeError("FastAPI 服务器启动超时")

    yield BASE_URL

    proc.terminate()
    proc.wait(timeout=10)


@pytest.fixture(autouse=True)
def clear_sources(server):
    """每个测试前清空服务器内存中的数据源，保证隔离性。"""
    httpx.delete(f"{server}/api/sources", timeout=10)
    yield
    httpx.delete(f"{server}/api/sources", timeout=10)
