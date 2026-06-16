#!/usr/bin/env python3
"""
环境检查脚本
=============
检查 Python 版本、依赖包、Excel 读取能力。
AI 在一切操作前调用此脚本，确保环境可用。

用法:
  python scripts/check_env.py --json        # 输出结构化结果
  python scripts/check_env.py --json --fix  # 自动安装缺失依赖

exit code: 0=OK, 1=有问题
"""

import argparse
import json
import re
import sys
import subprocess
from pathlib import Path


def check_python() -> dict:
    """检查 Python 版本"""
    v = sys.version_info
    ok = v >= (3, 10)
    return {
        "version": f"{v.major}.{v.minor}.{v.micro}",
        "ok": ok,
        "required": ">=3.10",
    }


def _parse_version(v: str) -> tuple:
    """将版本字符串解析为可比元组，不依赖 packaging 包。
    处理 2.0b1、2.3.3.post1、1.0.dev0 等特殊格式，只提取数字部分。"""
    parts = re.findall(r'\d+', str(v))[:3]
    return tuple(int(x) for x in parts) if parts else (0,)


def check_package(name: str, min_version: str = None) -> dict:
    """检查单个包是否安装及版本. min_version 传纯版本号如 '2.0'"""
    try:
        mod = __import__(name)
        installed = getattr(mod, '__version__', 'unknown')
        ok = True
        if min_version and installed != 'unknown':
            ok = _parse_version(str(installed)) >= _parse_version(min_version)
    except ImportError:
        installed = None
        ok = False
    except Exception:
        installed = str(getattr(mod, '__version__', 'unknown')) if 'mod' in dir() else 'unknown'
        ok = True  # 已安装但版本检查失败，不影响
    return {
        "installed": installed is not None,
        "version": str(installed) if installed else None,
        "min_version": min_version,
        "ok": ok,
    }


def check_excel_engines() -> dict:
    """检查可用的 Excel 读取引擎"""
    engines = {"polars_fastexcel": False, "pandas_openpyxl": False}
    try:
        import polars as pl
        pl.read_excel  # touch
        engines["polars_fastexcel"] = True
    except Exception:
        pass
    try:
        import openpyxl
        engines["pandas_openpyxl"] = True
    except Exception:
        pass
    return engines


def run_check() -> dict:
    """执行完整环境检查，返回结构化 dict"""
    pkgs = {
        "pandas": "2.0",
        "numpy": None,
        "scipy": None,
        "polars": "1.0",
        "openpyxl": None,
    }
    pkg_labels = {
        "pandas": ">=2.0",
        "numpy": "any",
        "scipy": "any",
        "polars": ">=1.0",
        "openpyxl": "any",
    }
    pkg_results = {k: check_package(k, v) for k, v in pkgs.items()}

    excel = check_excel_engines()
    can_read = excel.get("polars_fastexcel") or excel.get("pandas_openpyxl")

    missing = [k for k, v in pkg_results.items() if not v["ok"]]

    return {
        "python": check_python(),
        "packages": pkg_results,
        "excel_engines": excel,
        "can_read_excel": can_read,
        "missing_packages": missing,
        "all_ok": len(missing) == 0,
    }


def fix_missing():
    """安装 requirements.txt 中缺失的包"""
    req_path = Path(__file__).resolve().parent.parent / "requirements.txt"
    if not req_path.exists():
        print("[ERR] requirements.txt not found", file=sys.stderr)
        return False
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "-r", str(req_path)],
            stdout=sys.stdout, stderr=sys.stderr
        )
        return True
    except subprocess.CalledProcessError:
        return False


def main():
    parser = argparse.ArgumentParser(description="环境检查")
    parser.add_argument("--json", action="store_true", help="输出结构化 JSON")
    parser.add_argument("--fix", action="store_true", help="自动安装缺失的包")
    args = parser.parse_args()

    if args.fix:
        ok = fix_missing()
        sys.exit(0 if ok else 1)

    result = run_check()

    if args.json:
        print(json.dumps(result, ensure_ascii=False, indent=2))
    else:
        # 文本模式
        print(f"Python: {result['python']['version']} {'OK' if result['python']['ok'] else 'FAIL'}")
        for name, p in result['packages'].items():
            status = "OK" if p['ok'] else f"MISSING (need {p.get('min_version') or 'any'})"
            print(f"  {name}: {p['version'] or 'NOT INSTALLED'} {status}")
        print(f"Excel: polars_fastexcel={result['excel_engines']['polars_fastexcel']} openpyxl={result['excel_engines']['pandas_openpyxl']}")
        print(f"All OK: {result['all_ok']}")

    sys.exit(0 if result["all_ok"] else 1)


if __name__ == "__main__":
    main()
