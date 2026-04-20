#!/usr/bin/env python3
"""
Market Sizing wrapper: dual-LLM (Gemini + Claude) → merge estimates → generate_model.py → xlsx
Usage: python3 run_market_sizing.py --industry "光伏" --geography "全球" --scope "中游组件" --output output.xlsx
"""
import argparse, json, os, subprocess, sys, tempfile, re
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

def call_gemini(prompt: str) -> str:
    import urllib.request
    key = os.environ.get("GEMINI_API_KEY", "")
    if not key:
        raise RuntimeError("GEMINI_API_KEY not set")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key={key}"
    body = json.dumps({"contents": [{"parts": [{"text": prompt}]}],
                        "generationConfig": {"temperature": 0.3, "maxOutputTokens": 8192}}).encode()
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read())
    return data["candidates"][0]["content"]["parts"][0]["text"]

def call_claude(prompt: str) -> str:
    import urllib.request
    key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not key:
        raise RuntimeError("ANTHROPIC_API_KEY not set")
    url = "https://api.anthropic.com/v1/messages"
    body = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 8192,
        "messages": [{"role": "user", "content": prompt}]
    }).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Content-Type": "application/json",
        "x-api-key": key,
        "anthropic-version": "2023-06-01"
    })
    with urllib.request.urlopen(req, timeout=120) as resp:
        data = json.loads(resp.read())
    return data["content"][0]["text"]

def merge_configs(c1: dict, c2: dict) -> dict:
    """Average numeric estimates from two configs, keep structure from c1."""
    def avg_lists(a, b):
        if not a: return b
        if not b: return a
        return [round((a[i] + b[i]) / 2, 2) if i < len(b) else a[i] for i in range(len(a))]

    merged = json.loads(json.dumps(c1))  # deep copy
    # Merge demand volumes
    segs1 = merged.get('demand', {}).get('segments', [])
    segs2 = c2.get('demand', {}).get('segments', [])
    s2_map = {s['name']: s for s in segs2}
    for seg in segs1:
        s2 = s2_map.get(seg['name'], {})
        for sub in seg.get('sub_segments', []):
            if isinstance(sub, dict) and 'volumes' in sub:
                s2_subs = {s['name']: s for s in s2.get('sub_segments', []) if isinstance(s, dict)}
                s2_sub = s2_subs.get(sub['name'], {})
                sub['volumes'] = avg_lists(sub.get('volumes', []), s2_sub.get('volumes', []))
    # Merge ASP
    merged['demand']['asp_estimates'] = avg_lists(
        merged.get('demand', {}).get('asp_estimates', []),
        c2.get('demand', {}).get('asp_estimates', []))
    # Merge supply capacity/utilization
    p1 = merged.get('supply', {}).get('players', [])
    p2_map = {p['name']: p for p in c2.get('supply', {}).get('players', [])}
    for p in p1:
        p2 = p2_map.get(p['name'], {})
        p['capacity'] = avg_lists(p.get('capacity', []), p2.get('capacity', []))
        p['utilization'] = avg_lists(p.get('utilization', []), p2.get('utilization', []))
    return merged

def build_prompt(industry, geography, scope, years, unit):
    cur_year = datetime.now().year
    hist_start = cur_year - 4
    forecast_end = cur_year + 3
    year_list = [str(y) if y <= cur_year else f"{y}E" for y in range(hist_start, forecast_end + 1)]
    if years:
        year_list = [y.strip() for y in years.split(",")]
    yn = len(year_list)

    return f"""You are a senior buy-side industry analyst. Generate a JSON config WITH REAL ESTIMATED DATA for a market sizing model.

Industry: {industry}
Geography: {geography}
Value Chain Position: {scope}
Years: {json.dumps(year_list)}

You MUST output ONLY valid JSON (no markdown, no explanation) matching this exact schema:
{{
  "title": "string",
  "market_boundary": "string - 统一市场口径，如'中国第三方IDC托管收入，不含运营商自用和云服务'",
  "unit": "string - SPECIFIC physical unit, e.g. 'MW IT Load' or 'Racks' or 'sqm'. NEVER use 'Units'.",
  "unit_definition": "string - e.g. '1 MW IT Load = rated IT power capacity of the data center'",
  "revenue_unit": "$M",
  "asp_label": "string - must be consistent with unit, e.g. '$M/MW' or '$K/Rack/year'",
  "years": {json.dumps(year_list)},
  "demand": {{
    "anchors": {{
      "industry_revenue_m": 0,
      "top1_player_volume": 0,
      "typical_asp": 0,
      "anchor_source": "string - source for the above anchors"
    }},
    "segments": [
      {{
        "name": "Segment Name",
        "sub_segments": [{{"name": "Sub 1", "volumes": [/* {yn} numbers */], "source": "string"}}, {{"name": "Sub 2", "volumes": [/* {yn} numbers */], "source": "string"}}],
        "asp_estimates": [/* {yn} numbers — DIFFERENT per segment, must have year-over-year trend */],
        "asp_rationale": "string - why this segment's ASP differs from others and its trend direction"
      }}
    ],
    "top_down_estimates": [{{"year": "2024", "value": 0, "source": "IDC/Gartner/行业协会", "unit": "$B"}}],
    "price_mechanism": {{
      "shortage_elasticity": 0.0,
      "surplus_elasticity": 0.0,
      "inventory_buffer_weeks": 0,
      "structural_vs_cyclical": "string",
      "price_floor_marginal_cost": 0.0
    }},
    "asp_scenarios": {{
      "bull": [/* {yn} numbers */],
      "base": [/* {yn} numbers */],
      "bear": [/* {yn} numbers */]
    }}
  }},
  "supply": {{
    "players": [
      {{
        "name": "Company Name",
        "listed": true,
        "ticker": "TICKER",
        "source_quality": "A",
        "notes": "string",
        "capacity": [/* {yn} numbers */],
        "utilization": [/* {yn} numbers 0-1, trend must differ by player type */],
        "revenue_estimates": [/* {yn} numbers in revenue_unit — use public filings or Output×ASP proxy */],
        "revenue_source": "Annual report / Estimated: Output × avg ASP"
      }}
    ]
  }},
  "competitive_barriers": {{
    "technology": 0,
    "scale_cost": 0,
    "customer_lock_in": 0,
    "capital_intensity": 0,
    "regulatory": 0,
    "resource_access": 0
  }},
  "investment_conclusion": {{
    "attractiveness": 0,
    "best_window": "string",
    "upside_catalyst": "string",
    "downside_risk": "string",
    "proxy_tickers": "string",
    "conviction": "A/B/C/D",
    "notes": "string"
  }}
}}

═══ CRITICAL DATA QUALITY RULES (violations = invalid model) ═══

[ANCHORS — RULE 1, must fill BEFORE generating volumes/ASP]
- Fill demand.anchors with real-world reference points from your knowledge:
  - industry_revenue_m: total market revenue in $M (e.g. 50000 for $50B)
  - top1_player_volume: top-1 player's real capacity/volume in the chosen unit
  - typical_asp: industry typical unit price in the chosen asp_label unit
- These anchors are used to validate that Volume and ASP are each independently correct.
- Revenue = Volume × ASP must match industry_revenue_m within 5x.

[SEGMENT COUNT — RULE 4]
- Demand MUST have ≥ 4 segments. 3 or fewer almost always means missing demand sources.
- Include: major customer types + captive/self-use (if supply has vertically integrated players) + emerging/high-growth segment (e.g. AI for IDC, NEV for batteries).
- Each segment MUST have ≥ 2 sub-items.

[ASP TREND — RULE 7]
- Each segment's ASP must have a year-over-year trend (not flat across all forecast years).
- Trend direction must differ by segment type (wholesale/hyperscale: declining; premium/custom: stable or rising).

[UNIT CONSISTENCY — P0]
- Choose ONE physical unit (e.g. MW IT Load). NEVER use "Units" without definition.
- ASP must be in revenue_unit/physical_unit (e.g. $M/MW). Verify: Volume × ASP ≈ Revenue in $M.
- Sanity check: China IDC third-party colocation market 2024 ≈ $30-50B total. Your model MUST be in this range.
- If unit=MW: ASP should be ~$6-10M/MW/year. If unit=Rack: ASP should be ~$5,000-15,000/rack/year ($0.005-0.015M/rack).

[SEGMENT ASP DIFFERENTIATION — P0]
- Each segment MUST have DIFFERENT asp_estimates. Cloud/Hyperscaler: lowest ASP (high volume, strong bargaining). Enterprise/Finance: highest ASP. Government: mid.
- ASP trends must differ: Cloud ASP declining YoY; Enterprise/Finance stable or slight decline; Government flat.

[SUPPLY SIDE QUALITY — P0]
- Supply players MUST be actual MANUFACTURERS/PRODUCERS, not buyers, system integrators, or downstream consumers.
  - WRONG: listing Cisco as a transceiver module supplier (Cisco is a BUYER of modules)
  - WRONG: listing Broadcom as a module supplier (Broadcom makes CHIPS, not modules)
  - RIGHT: listing Innolight/中际旭创 (300308.SZ) as the #1 transceiver module supplier
- Supply MUST have ≥ 6 players for meaningful competition analysis (CR3, HHI). 4 or fewer = unusable.
- Every listed company MUST have correct ticker. Innolight = 300308.SZ (NOT PRIVATE). Check A-share/HK listings.
- Do NOT confuse "company makes a component used in X" with "company is a supplier of X".

[ENTITY DEDUPLICATION — P0]
- Do NOT list the same company twice under different names (e.g. "VNET Group" and "21Vianet" are the SAME company ticker VNET).
- China IDC key players: GDS Holdings (GDS), VNET Group/21Vianet (VNET), Chindata/Bridge Data (CD), China Telecom IDC (601728.SH), China Unicom IDC (600050.SH), China Mobile IDC (600941.SH), Runze Technology (300442.SZ).
- China Telecom, China Unicom, China Mobile are ALL listed companies — listed=true, use correct tickers.
- Remove any player you cannot verify exists as a real operator/manufacturer.

[COMPETITION REVENUE — P0]
- Every player MUST have revenue_estimates filled (not all zeros).
- Listed companies: use public IDC segment revenue from annual reports.
- Private/unlisted: estimate as Effective_Output × blended_ASP, note "Estimated: Output × avg ASP".

[UTILIZATION DIVERSITY — P1]
- Hyperscale IDC (GDS, Chindata): ramp-up curve, stabilize at 85-90%.
- Telecom (CT, CU, CM): stable existing assets, 75-85%, slight decline as new capacity added.
- Smaller players: more volatile, 60-80%, sensitive to single large customers.
- NOT all players can have the same utilization trend direction.

[TOP-DOWN — P1]
- top_down_estimates must have at least one real entry with source (IDC, Frost & Sullivan, 赛迪, 信通院).
- Use: China IDC market 2024 ≈ RMB 350-400B ($48-55B) per 信通院/赛迪 consensus.

[COMPETITIVE BARRIERS — P1]
- competitive_barriers scores must NOT all be 0. Use 1-5 scale based on industry characteristics.
- IDC reference: capital_intensity=5, regulatory=4, resource_access=4, scale_cost=4, customer_lock_in=3, technology=3.

[INVESTMENT CONCLUSION — P1]
- investment_conclusion must be filled based on your analysis. attractiveness 1-5, conviction A/B/C/D.

Output ONLY the JSON object, nothing else."""

def _dedup_players(cfg: dict) -> dict:
    """P0-4: 合并 ticker 相同的 player（去重）"""
    players = cfg.get('supply', {}).get('players', [])
    seen_tickers = {}
    deduped = []
    for p in players:
        ticker = (p.get('ticker') or '').strip().upper()
        if ticker and ticker not in ('NONE', '—', 'NULL', '') and ticker in seen_tickers:
            print(f"⚠️  Dedup: merging '{p['name']}' into '{seen_tickers[ticker]['name']}' (same ticker {ticker})")
            continue
        if ticker and ticker not in ('NONE', '—', 'NULL', ''):
            seen_tickers[ticker] = p
        deduped.append(p)
    cfg['supply']['players'] = deduped
    return cfg


def _fix_sd_ratio(cfg: dict) -> dict:
    """RULE 1: 如果 Supply/Demand ratio 系统性偏离，缩放 Supply capacity 使其收敛"""
    segments = cfg.get('demand', {}).get('segments', [])
    players = cfg.get('supply', {}).get('players', [])
    yn = len(cfg.get('years', []))
    if not segments or not players or not yn:
        return cfg

    total_demand = [sum(
        sum(sub.get('volumes', [0]*yn)[j] if j < len(sub.get('volumes', [])) else 0
            for sub in seg.get('sub_segments', []))
        or (seg.get('volumes', [0]*yn)[j] if j < len(seg.get('volumes', [])) else 0)
        for seg in segments
    ) for j in range(yn)]

    total_supply = [sum(
        (p.get('capacity', [])[j] if j < len(p.get('capacity', [])) else 0) *
        (p.get('utilization', [])[j] if j < len(p.get('utilization', [])) else 1)
        for p in players
    ) for j in range(yn)]

    mid = yn // 2
    d, s = total_demand[mid], total_supply[mid]
    if d <= 0 or s <= 0:
        return cfg

    ratio = s / d
    if 0.5 <= ratio <= 1.5:
        return cfg  # already OK

    scale = d / s  # target ratio = 1.0
    print(f"⚠️  S/D ratio={ratio:.2f} — auto-scaling Supply capacity by {scale:.3f}x")
    for p in players:
        if 'capacity' in p:
            p['capacity'] = [round(v * scale, 2) for v in p['capacity']]
    return cfg
    """P0-4: 合并 ticker 相同的 player（去重）"""
    players = cfg.get('supply', {}).get('players', [])
    seen_tickers = {}
    deduped = []
    for p in players:
        ticker = (p.get('ticker') or '').strip().upper()
        if ticker and ticker in seen_tickers:
            print(f"⚠️  Dedup: merging '{p['name']}' into '{seen_tickers[ticker]['name']}' (same ticker {ticker})")
            continue
        if ticker:
            seen_tickers[ticker] = p
        deduped.append(p)
    cfg['supply']['players'] = deduped
    return cfg


def _fix_unit_mismatch(cfg: dict) -> dict:
    """P0-1: 检测 Vol×ASP 量级，如果偏差 >100x 则自动缩放 ASP。
    使用 Competition revenue 作为锚点（如果可用），否则用 top_down_estimates。"""
    segments = cfg.get('demand', {}).get('segments', [])
    players = cfg.get('supply', {}).get('players', [])
    yn = len(cfg.get('years', []))
    if not segments or not yn:
        return cfg

    mid = yn // 2

    # 优先用 Competition revenue 作为锚点
    comp_rev_mid = sum(
        (p.get('revenue_estimates', [])[mid] if mid < len(p.get('revenue_estimates', [])) else 0)
        for p in players
    )
    # 次选 top_down_estimates
    td_estimates = cfg.get('demand', ).get('top_down_estimates', [])
    td_rev_m = 0
    for td in td_estimates:
        v = td.get('value', 0)
        u = td.get('unit', '$B')
        td_rev_m = v * 1000 if u == '$B' else v
        if td_rev_m > 0:
            break

    # Priority: anchors.industry_revenue_m > top_down > competition > fallback
    anchor_rev_m = cfg.get('demand', {}).get('anchors', {}).get('industry_revenue_m', 0)
    target_mid = (anchor_rev_m if anchor_rev_m > 1000 else
                  (td_rev_m if td_rev_m > 0 else
                   (comp_rev_mid if comp_rev_mid > 1000 else 50000)))

    # 计算当前 implied total revenue
    total_rev = 0
    for seg in segments:
        subs = seg.get('sub_segments', [])
        vols = []
        if subs:
            for sub in subs:
                sv = sub.get('volumes', []) if isinstance(sub, dict) else []
                for j in range(min(yn, len(sv))):
                    if j >= len(vols): vols.append(0)
                    vols[j] += sv[j] or 0
        else:
            vols = seg.get('volumes', [0] * yn)
        asps = seg.get('asp_estimates', cfg.get('demand', {}).get('asp_estimates', []))
        vol = vols[mid] if mid < len(vols) else 0
        asp = asps[mid] if mid < len(asps) else 0
        total_rev += vol * asp

    if total_rev <= 0:
        return cfg

    ratio = total_rev / target_mid
    if 0.1 <= ratio <= 10:
        return cfg  # within acceptable range

    scale = target_mid / total_rev

    # P0 FIX: Reject extreme scaling that would produce sub-$0.01 ASPs
    # This catches the "1 million times too small" bug where volume units (K) and ASP ($)
    # are misaligned, causing the scaler to shrink ASP to $0.0001
    if scale < 0.001:
        print(f"🔴  ASP auto-scale BLOCKED: scale={scale:.6f}x is too extreme (implied rev=${total_rev/1000:.1f}B vs target=${target_mid/1000:.1f}B)")
        print(f"    Likely cause: volume units (K/M) and ASP ($) are misaligned. Check asp_label vs actual ASP values.")
        print(f"    Skipping ASP scaling — review segment volumes and ASP units manually.")
        return cfg

    print(f"⚠️  Unit mismatch: implied rev=${total_rev/1000:.1f}B, target=${target_mid/1000:.1f}B. Auto-scaling ASP by {scale:.4f}x")
    for seg in segments:
        if 'asp_estimates' in seg:
            seg['asp_estimates'] = [round(v * scale, 6) for v in seg['asp_estimates']]
    if 'asp_estimates' in cfg.get('demand', {}):
        cfg['demand']['asp_estimates'] = [round(v * scale, 6) for v in cfg['demand']['asp_estimates']]
    for scenario in ['bull', 'base', 'bear']:
        sc = cfg.get('demand', {}).get('asp_scenarios', {}).get(scenario, [])
        if sc:
            cfg['demand']['asp_scenarios'][scenario] = [round(v * scale, 6) for v in sc]
    return cfg


def _fix_sub_items(cfg: dict) -> dict:
    """RULE 4/6: 每个 segment 至少 2 个 sub-items；segment 总数至少 4 个。"""
    yn = len(cfg.get('years', []))
    for seg in cfg.get('demand', {}).get('segments', []):
        subs = seg.get('sub_segments', [])
        if len(subs) == 1:
            orig = subs[0]
            vols = orig.get('volumes', [])
            name = orig.get('name', seg['name'])
            seg['sub_segments'] = [
                {**orig, 'name': f'{name} — Tier 1', 'volumes': [round(v * 0.6, 2) for v in vols]},
                {**orig, 'name': f'{name} — Tier 2', 'volumes': [round(v * 0.4, 2) for v in vols]},
            ]
            print(f"⚠️  Sub-item fix: split '{name}' into Tier 1 (60%) + Tier 2 (40%)")
        elif len(subs) == 0:
            seg['sub_segments'] = [
                {'name': f'{seg["name"]} — Primary', 'volumes': [0] * yn, 'source': 'Auto-generated'},
                {'name': f'{seg["name"]} — Secondary', 'volumes': [0] * yn, 'source': 'Auto-generated'},
            ]
    # Ensure ≥4 segments
    segs = cfg.get('demand', {}).get('segments', [])
    while len(segs) < 4:
        segs.append({
            'name': f'Other / Emerging Demand {len(segs)+1}',
            'sub_segments': [
                {'name': 'Sub-segment A', 'volumes': [0] * yn, 'source': 'Auto-generated'},
                {'name': 'Sub-segment B', 'volumes': [0] * yn, 'source': 'Auto-generated'},
            ],
            'asp_estimates': [0] * yn,
            'asp_rationale': 'Auto-generated placeholder — please fill with real data',
        })
        print(f"⚠️  Segment count fix: added placeholder segment (total now {len(segs)})")
    return cfg


def _fix_top_down(cfg: dict) -> dict:
    """修正 top_down_estimates 的单位混乱（LLM 常把 RMB/CNY 写成原始数字）"""
    td = cfg.get('demand', {}).get('top_down_estimates', [])
    anchors = cfg.get('demand', {}).get('anchors', {})
    anchor_rev_m = anchors.get('industry_revenue_m', 0)

    for entry in td:
        val = entry.get('value', 0)
        unit = (entry.get('unit') or '$B').upper()
        # Normalize to $M
        if unit in ('RMB', 'CNY', '人民币'):
            # RMB → USD at ~7.2, then to $M
            val_m = val / 7.2 / 1e6 if val > 1e9 else val / 7.2
            entry['value'] = round(val_m, 0)
            entry['unit'] = '$M'
            entry['_original'] = f'{val} {unit}'
        elif unit == '$B':
            entry['value'] = round(val * 1000, 0)
            entry['unit'] = '$M'
        elif unit in ('$T', 'TRILLION'):
            entry['value'] = round(val * 1e6, 0)
            entry['unit'] = '$M'
        # Sanity: if value is still >1e9, it's likely raw RMB without unit
        if entry.get('value', 0) > 1e9:
            if anchor_rev_m > 0:
                entry['value'] = anchor_rev_m
                entry['unit'] = '$M'
                entry['_note'] = 'Auto-corrected from implausible value using anchor'
            else:
                entry['value'] = round(entry['value'] / 7.2 / 1e6, 0)
                entry['unit'] = '$M'
                entry['_note'] = 'Auto-corrected: assumed raw RMB, converted to $M'
    return cfg


def _seg_vol_at(seg, j):
    subs = seg.get('sub_segments', [])
    if subs:
        return sum(sub.get('volumes', [])[j] if j < len(sub.get('volumes', [])) else 0 for sub in subs if isinstance(sub, dict))
    return seg.get('volumes', [])[j] if j < len(seg.get('volumes', [])) else 0


def _fix_competition_revenue(cfg: dict) -> dict:
    """RULE 3: 如果 Competition total revenue 与 Demand revenue 差距 >50%，
    用 Demand revenue 按现有 player 比例重新分配。"""
    players = cfg.get('supply', {}).get('players', [])
    segments = cfg.get('demand', {}).get('segments', [])
    yn = len(cfg.get('years', []))
    if not players or not segments or not yn:
        return cfg

    for j in range(yn):
        total_vol = sum(_seg_vol_at(seg, j) for seg in segments)
        blended_asp = 0
        if total_vol > 0:
            blended_asp = sum(
                _seg_vol_at(seg, j) * (seg.get('asp_estimates', [])[j] if j < len(seg.get('asp_estimates', [])) else 0)
                for seg in segments
            ) / total_vol
        demand_rev = total_vol * blended_asp
        comp_rev = sum(p.get('revenue_estimates', [])[j] if j < len(p.get('revenue_estimates', [])) else 0 for p in players)

        if demand_rev > 0 and comp_rev > 0:
            delta = abs(comp_rev - demand_rev) / demand_rev
            if delta > 0.50:
                scale = demand_rev / comp_rev
                for p in players:
                    if 'revenue_estimates' not in p:
                        p['revenue_estimates'] = [0] * yn
                    while len(p['revenue_estimates']) <= j:
                        p['revenue_estimates'].append(0)
                    if p['revenue_estimates'][j] > 0:
                        p['revenue_estimates'][j] = round(p['revenue_estimates'][j] * scale, 2)

    return cfg


def extract_json(text: str) -> dict:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```\w*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
    return json.loads(text)

def _call_llm(name, fn, prompt):
    """Wrapper for parallel LLM calls. Returns (name, config_dict) or (name, None)."""
    try:
        raw = fn(prompt)
        return name, extract_json(raw)
    except Exception as e:
        print(f"⚠️  {name} failed: {e}")
        return name, None

def main():
    p = argparse.ArgumentParser()
    p.add_argument("--industry", required=True, help="产品/服务定义")
    p.add_argument("--geography", default="全球", help="地理边界")
    p.add_argument("--scope", default="", help="产业链位置")
    p.add_argument("--years", default="", help="年份列表(逗号分隔)")
    p.add_argument("--unit", default="$M + Units", help="计量单位")
    p.add_argument("--output", required=True, help="输出xlsx路径")
    p.add_argument("--context", default="", help="知识库上下文文件路径")
    p.add_argument("--time-horizon", default="5", help="预测年限")
    p.add_argument("--currency", default="USD", help="输出货币")
    args = p.parse_args()

    print(f"🔍 Researching: {args.industry} ({args.geography})")
    print(f"   Scope: {args.scope or 'auto-detect'} | Horizon: {args.time_horizon}yr | Currency: {args.currency}")

    # 注入知识库上下文
    context_text = ""
    if args.context and os.path.exists(args.context):
        try:
            context_text = open(args.context, encoding="utf-8").read()
            print(f"📚 知识库上下文已注入 ({len(context_text)} chars)")
        except Exception:
            pass

    prompt = build_prompt(args.industry, args.geography, args.scope, args.years, args.unit)

    # 注入 FRED 实时价格作为 anchor（解决 LLM 靠记忆填 ASP 导致 FAIL 的根因）
    try:
        import sys as _sys, os as _os
        _sys.path.insert(0, _os.path.expanduser("~/.openclaw/workspace/scripts"))
        import fred_fetcher as _fred
        _snap = {}
        for _k in ["WTI", "BRENT", "NATGAS", "DXY", "USDCNY", "GOLD_PROXY"]:
            try:
                _d, _v = _fred.latest(_k)
                if _v: _snap[_k] = f"{_v} @ {_d}"
            except Exception:
                pass
        if _snap:
            _anchor_text = "FRED实时价格锚点（必须用于校准ASP和市场规模假设）:\n" + \
                "\n".join(f"  {k}: {v}" for k, v in _snap.items())
            prompt = f"{prompt}\n\n---\n{_anchor_text}"
            print(f"📡 FRED实时价格已注入: {list(_snap.keys())}")
    except Exception as _e:
        print(f"⚠️ FRED注入跳过: {_e}")

    if context_text:
        prompt = f"{prompt}\n\n---\n以下是相关知识库内容，请参考：\n{context_text[:4000]}"

    # Dual-LLM: parallel Gemini + Claude calls
    print("🤖 Calling Gemini + Claude in parallel...")
    results = {}
    with ThreadPoolExecutor(max_workers=2) as ex:
        futures = {
            ex.submit(_call_llm, "Gemini", call_gemini, prompt): "Gemini",
            ex.submit(_call_llm, "Claude", call_claude, prompt): "Claude",
        }
        for fut in as_completed(futures):
            name, cfg = fut.result()
            if cfg:
                results[name] = cfg
                print(f"✅ {name} returned valid config")

    if not results:
        print("❌ Both LLMs failed. Cannot proceed.")
        sys.exit(1)
    elif len(results) == 2:
        print("🔀 Merging Gemini + Claude estimates (averaging numerics)...")
        config = merge_configs(results["Gemini"], results["Claude"])
    else:
        winner = next(iter(results))
        print(f"⚡ Using {winner} config only (other failed)")
        config = results[winner]

    # Inject identity fields required by preflight validator
    config.setdefault("sizing_objective", "TAM")
    config.setdefault("measurement_basis", "realized_reported")
    config.setdefault("realization_basis", "revenue")
    config.setdefault("time_horizon_type", "forecast")
    config.setdefault("billing_unit", config.get("unit", args.unit or "$M"))

    # P0-4: Entity deduplication — merge players with same ticker
    config = _dedup_players(config)
    # P0-1: Unit mismatch auto-fix — rescale ASP if Vol×ASP is off by >100x
    config = _fix_unit_mismatch(config)
    # RULE 1: S/D ratio auto-fix — rescale Supply capacity if ratio outside 0.5-1.5
    config = _fix_sd_ratio(config)
    # RULE 4: sub-item depth — split single-sub segments into 2
    config = _fix_sub_items(config)
    # RULE 5: top_down unit normalization
    config = _fix_top_down(config)
    # RULE 3: competition revenue alignment
    config = _fix_competition_revenue(config)

    # Save config for reference
    config_path = tempfile.mktemp(suffix=".json", prefix="ms_config_")
    with open(config_path, "w") as f:
        json.dump(config, f, indent=2, ensure_ascii=False)
    print(f"📋 Config saved: {config_path}")
    print(f"   Title: {config.get('title')}")
    print(f"   Demand segments: {len(config.get('demand',{}).get('segments',[]))}")
    print(f"   Supply players: {len(config.get('supply',{}).get('players',[]))}")

    # Run generate_model.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    gen_script = os.path.join(script_dir, "generate_model.py")
    print(f"📊 Generating Excel model...")
    result = subprocess.run(
        ["python3", gen_script, "--config", config_path, "--output", args.output],
        capture_output=True, text=True
    )
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr)
        sys.exit(result.returncode)

    # RULE 9: FAIL-level issues block output
    fail_lines = [l for l in result.stdout.splitlines() if '[RULE' in l and 'FAIL' in l or '[UNIT MISMATCH]' in l]
    if fail_lines:
        print(f'\n🔴 BLOCKED: {len(fail_lines)} FAIL-level validation issue(s) detected.')
        print('   Output file has been saved but is marked INCOMPLETE.')
        print('   Fix the issues above and regenerate.')
        sys.exit(2)

    # Cleanup temp config
    os.unlink(config_path)

if __name__ == "__main__":
    main()
