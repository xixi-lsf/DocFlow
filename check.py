#!/usr/bin/env python3
"""
Project Completeness Check Script
Verify all necessary files and configurations are ready
"""

import os
import sys
from pathlib import Path

def check_file(path, description):
    """Check if file exists"""
    exists = Path(path).exists()
    status = "[OK]" if exists else "[FAIL]"
    print(f"  {status} {description}: {path}")
    return exists

def check_directory(path, description):
    """Check if directory exists"""
    exists = Path(path).is_dir()
    status = "[OK]" if exists else "[FAIL]"
    print(f"  {status} {description}: {path}")
    return exists

def main():
    print("\n" + "="*60)
    print("  Document Filler System - Project Check")
    print("="*60 + "\n")

    all_ok = True

    # Check core code files
    print("Core Code Files:")
    core_files = [
        ("backend/main.py", "FastAPI main entry"),
        ("backend/agent.py", "AI data extraction"),
        ("backend/extractor.py", "Document parsing"),
        ("backend/filler.py", "Table filling"),
        ("frontend/index.html", "Web interface"),
        ("frontend/app.js", "Frontend interaction"),
    ]
    for path, desc in core_files:
        if not check_file(path, desc):
            all_ok = False

    # Check config files
    print("\nConfiguration Files:")
    config_files = [
        ("requirements.txt", "Python dependencies"),
        ("config.ini", "Project configuration"),
        (".env.example", "Environment variables example"),
    ]
    for path, desc in config_files:
        if not check_file(path, desc):
            all_ok = False

    # Check startup scripts
    print("\nStartup Scripts:")
    script_files = [
        ("run.bat", "Windows startup script"),
        ("run.sh", "Linux/Mac startup script"),
    ]
    for path, desc in script_files:
        if not check_file(path, desc):
            all_ok = False

    # Check documentation
    print("\nDocumentation Files:")
    doc_files = [
        ("README.md", "Project description"),
        ("QUICKSTART.md", "Quick start guide"),
        ("DEPLOYMENT.md", "Deployment guide"),
        ("TESTING.md", "Testing guide"),
        ("PROJECT_SUMMARY.md", "Project summary"),
    ]
    for path, desc in doc_files:
        if not check_file(path, desc):
            all_ok = False

    # Check directories
    print("\nProject Directories:")
    dirs = [
        ("backend", "Backend directory"),
        ("frontend", "Frontend directory"),
        ("outputs", "Output directory"),
    ]
    for path, desc in dirs:
        if not check_directory(path, desc):
            all_ok = False

    # Check virtual environment
    print("\nPython Environment:")
    venv_exists = Path(".venv").is_dir()
    status = "[OK]" if venv_exists else "[INFO]"
    print(f"  {status} Virtual environment: .venv")
    if not venv_exists:
        print("     (Will be created automatically on first run)")

    # Check API Key
    print("\nAPI Configuration:")
    api_key = os.environ.get("DEEPSEEK_API_KEY", "")
    if api_key:
        masked_key = api_key[:10] + "..." + api_key[-4:]
        print(f"  [OK] DEEPSEEK_API_KEY is set: {masked_key}")
    else:
        print(f"  [FAIL] DEEPSEEK_API_KEY is not set")
        print("     Please run: set DEEPSEEK_API_KEY=sk-your-api-key (Windows)")
        print("     Or run: export DEEPSEEK_API_KEY=sk-your-api-key (Linux/Mac)")
        all_ok = False

    # Check dependencies
    print("\nPython Dependencies:")
    try:
        import fastapi
        print(f"  [OK] FastAPI: {fastapi.__version__}")
    except ImportError:
        print(f"  [FAIL] FastAPI not installed")
        all_ok = False

    try:
        import uvicorn
        print(f"  [OK] Uvicorn installed")
    except ImportError:
        print(f"  [FAIL] Uvicorn not installed")
        all_ok = False

    try:
        import docx
        print(f"  [OK] python-docx installed")
    except ImportError:
        print(f"  [FAIL] python-docx not installed")
        all_ok = False

    try:
        import openpyxl
        print(f"  [OK] openpyxl installed")
    except ImportError:
        print(f"  [FAIL] openpyxl not installed")
        all_ok = False

    # Final result
    print("\n" + "="*60)
    if all_ok:
        print("SUCCESS: All checks passed! Project is ready")
        print("\nNext steps:")
        print("  1. Make sure DEEPSEEK_API_KEY environment variable is set")
        print("  2. Run startup script:")
        print("     Windows: run.bat")
        print("     Linux/Mac: ./run.sh")
        print("  3. Open browser to http://localhost:8000")
        return 0
    else:
        print("FAILED: Some checks failed, please fix the issues above")
        print("\nCommon issues:")
        print("  - Missing files: Check project structure is complete")
        print("  - Missing dependencies: Run pip install -r requirements.txt")
        print("  - Missing API Key: Set DEEPSEEK_API_KEY environment variable")
        return 1

if __name__ == "__main__":
    sys.exit(main())
