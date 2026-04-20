#!/usr/bin/env python3
"""
Thin wrapper for the 产业链 Screener — called by OpenClaw skill.
Usage: python3 run_chain_screener.py "SpaceX产业链" [--output ./screener_output]
"""

import os
import sys
from pathlib import Path

quant_root = os.getenv("QUANT_ROOT")
if not quant_root:
    print("❌ QUANT_ROOT not set. Example: export QUANT_ROOT=\"/Users/harryhuang/Algo Trading/Quant Trading\"")
    sys.exit(1)

sys.path.insert(0, quant_root)

from dotenv import load_dotenv
load_dotenv(Path(quant_root) / ".env")
load_dotenv(Path(quant_root) / "configs" / ".env")

import argparse
import json
import tushare as ts
from src.ai.router import LLMRouter
from src.data_loader.fmp_stable import FMPStableClient
from src.analysis.chain_screener import run_screener


def _init_ts():
    pro = ts.pro_api("init")
    token = os.getenv("TUSHARE_TOKEN", "")
    api_url = os.getenv("TUSHARE_API_URL", "")
    if token:
        pro._DataApi__token = token
    if api_url:
        pro._DataApi__http_url = api_url
    return pro


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("theme", help="投资主题，如 'SpaceX产业链' / 'AI算力' / '人形机器人'")
    parser.add_argument("--output", default=os.path.join(quant_root, "screener_output"))
    args = parser.parse_args()

    summary = run_screener(
        theme=args.theme,
        llm_router=LLMRouter(),
        fmp_client=FMPStableClient(),
        ts_pro=_init_ts(),
        output_dir=args.output,
    )
    print("\nSUMMARY_JSON:" + json.dumps(summary, ensure_ascii=False))


if __name__ == "__main__":
    main()
