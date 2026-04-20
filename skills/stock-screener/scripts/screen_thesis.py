#!/usr/bin/env python3
"""
OpenClaw skill wrapper for stock screener v2.
Calls the FastAPI endpoint to perform thematic stock screening.
"""
import argparse
import json
import os
import sys
import requests
from typing import Optional

SCREENER_API_URL = "http://localhost:8000"

def check_service_running() -> bool:
    """Check if the screener service is running."""
    try:
        response = requests.get(f"{SCREENER_API_URL}/health", timeout=2)
        return response.status_code == 200
    except:
        return False

def screen_thesis(thesis: str, top_n: int = 10, aggression: Optional[float] = None) -> dict:
    """Call the screener API to screen stocks by thesis."""
    payload = {
        "thesis": thesis,
        "max_results": top_n,
        "aggression_override": aggression
    }
    
    response = requests.post(
        f"{SCREENER_API_URL}/screen",
        json=payload,
        timeout=300  # 5 minutes timeout
    )
    response.raise_for_status()
    return response.json()

def format_output(result: dict) -> str:
    """Format the screening result for display."""
    lines = []
    
    # Thesis validation
    validation = result["thesis_validation"]
    lines.append(f"✅ 主题评分：{validation['score']}/10")
    lines.append(f"📝 判断：{validation['verdict']}")
    if validation.get('evidence_for'):
        lines.append(f"💡 支持证据：{', '.join(validation['evidence_for'][:2])}")
    lines.append("")
    
    # Discovery
    lines.append(f"🔍 发现 {result['candidates_found']} 个相关候选股票")
    lines.append("")
    
    # Top picks
    if not result["top_picks"]:
        lines.append("❌ 未找到符合条件的股票")
        return "\n".join(lines)
    
    lines.append(f"🏆 前 {len(result['top_picks'])} 个推荐股票：")
    lines.append("")
    
    for i, pick in enumerate(result["top_picks"], 1):
        factors = pick["factor_breakdown"]
        smart = pick["smart_money_signals"]
        
        lines.append(f"【{i}】{pick['ticker']} - {pick['name']}")
        lines.append(f"  💰 市值：${pick['market_cap']/1e9:.1f}B")
        lines.append(f"  📊 综合评分：{pick['composite_score']:.1f} | 量化：{pick['quant_score']:.1f} | 聪明钱：{pick['smart_money_score']:.1f}")
        
        # Factor breakdown
        stars = lambda x: "★" * int(x/2) + "☆" * (5 - int(x/2))
        lines.append(f"  📈 因子：估值 {stars(factors['valuation'])} | 成长 {stars(factors['growth'])} | 动量 {stars(factors['momentum'])} | 质量 {stars(factors['quality'])}")
        
        # Technical & Money Flow Summary
        if pick.get("technical_summary"):
            lines.append(f"  📉 技术面：{pick['technical_summary']}")
        
        lines.append(f"  💡 {pick['summary']}")
        lines.append("")
    
    return "\n".join(lines)

def main():
    parser = argparse.ArgumentParser(description="Screen stocks by investment thesis")
    parser.add_argument("thesis", help="Investment thesis (Chinese or English)")
    parser.add_argument("--top", type=int, default=10, help="Number of top picks (default: 10)")
    parser.add_argument("--aggression", type=float, help="Aggression level (0.5-2.0)")
    
    args = parser.parse_args()
    
    # Validate arguments
    if args.top < 1 or args.top > 30:
        print("❌ Error: --top must be between 1 and 30", file=sys.stderr)
        sys.exit(1)
    
    if args.aggression and (args.aggression < 0.5 or args.aggression > 2.0):
        print("❌ Error: --aggression must be between 0.5 and 2.0", file=sys.stderr)
        sys.exit(1)
    
    # Check if service is running
    if not check_service_running():
        print("❌ Screener 服务未运行", file=sys.stderr)
        print("", file=sys.stderr)
        print("请先启动服务：", file=sys.stderr)
        print(f"  cd $SCREENER_ROOT", file=sys.stderr)
        print(f"  source venv/bin/activate", file=sys.stderr)
        print(f"  uvicorn main:app --host 0.0.0.0 --port 8000", file=sys.stderr)
        sys.exit(1)
    
    # Run screening
    print(f"⏳ 正在分析「{args.thesis}」主题，预计需要1-3分钟...")
    print("")
    
    try:
        result = screen_thesis(args.thesis, args.top, args.aggression)
        output = format_output(result)
        print(output)
        
        # Disclaimer
        print("")
        print("⚠️ 免责声明：本筛选结果仅供参考，不构成投资建议。投资有风险，决策需谨慎。")
        
    except requests.exceptions.RequestException as e:
        print(f"❌ API 调用失败：{e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"❌ 执行失败：{e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
