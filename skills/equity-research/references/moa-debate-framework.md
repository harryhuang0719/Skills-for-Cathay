# MoA Debate Framework (Mixture of Agents)

Deep mode only. Claude performs each agent role sequentially, maintaining clear labels for each agent's output. All analysis output in Chinese (简体中文).

**Currency Rule**: All price outputs use the same currency as input data. US stocks = USD, A-shares = CNY, HK stocks = HKD. Never convert currencies.

---

## Execution Flow

```
Phase 1 (Foundation):  Macro Strategist → Forensic Accountant → Industry Specialist
Phase 2 (Advocacy):    Bull Advocate → Bear Advocate
Phase 2.5 (Rebuttal):  Bull Rebuttal → Bear Rebuttal
Phase 3 (Decision):    CIO (independent view FIRST, then evaluate agents)
```

---

## Phase 1: Foundation Agents

### 1a. Macro Strategist (宏观策略师)

Analyze the macro environment relevant to this company:
- Current economic cycle position (expansion/peak/contraction/trough)
- Interest rate environment and trajectory
- Sector rotation signals — is money flowing into or out of this sector?
- Currency and commodity price impacts
- Geopolitical risks affecting this company/sector
- Regime assessment: calm / normal / stress

Output: 300-500 words, focused conclusions with supporting data.

### 1b. Forensic Accountant (法务会计师)

Scrutinize financial quality:
- Earnings quality: accruals ratio, cash conversion (OCF/NI should be >1.0)
- Revenue recognition: any aggressive practices? Channel stuffing signals?
- Cash flow analysis: FCF trend vs net income trend — divergence is a red flag
- Balance sheet health: leverage ratios, off-balance-sheet items, goodwill/intangibles ratio
- Working capital trends: DSO/DIO/DPO changes
- Related party transactions, auditor changes, restatements
- Red flags: list any concerns with severity (high/medium/low)

Output: 300-500 words. Conclude with a financial quality score (1-10, where 10 = pristine).

### 1c. Industry Specialist (行业专家)

Use the paradigm-specific prompt fragment from `valuation-paradigms.md` for the matched industry. Additionally:
- Industry structure: fragmented vs consolidated, barriers to entry
- Competitive positioning: market share, moat type and durability
- Industry lifecycle: growth/maturity/decline
- Structural changes: technology disruption, regulation, supply chain shifts
- Key success factors in this industry

Output: 300-500 words. Conclude with competitive advantage assessment (strong/moderate/weak/none).

---

## Phase 2: Advocacy

### 2a. Bull Advocate (做多基金经理)

**Identity**: Aggressive long-only fund manager. 32% historical annualized return. Finds market mispricing, concentrates bets, earns returns from information asymmetry.

**Stance**: Build the strongest possible long case for {TICKER}. If no strong case exists, conclude "不建议做多" (conviction <30).

**Required Analysis** (complete every item):

**1. Core Thesis (3 sentences)**:
- What to buy? (增长加速 / 估值修复 / 周期反转 / 事件驱动 / 资产重估)
- Market mispricing? (what wrong expectation is priced in?)
- When does it pay off? (specific catalyst + time window)

**2. Catalyst Timeline** (minimum 3):

| Catalyst | Expected Date | Impact Quantified | Probability |
|----------|--------------|-------------------|-------------|
| e.g., Q1 earnings beat | 2026-04-25 | EPS beat 10%+ → stock +8-12% | 65% |

No date = not a catalyst. Must be specific.

**3. Price Targets** (3 scenarios with valuation basis):

| Scenario | Target | Method & Calculation | Probability | Timeframe |
|----------|--------|---------------------|-------------|-----------|
| Base | $XXX | Based on [method]: [calculation] | XX% | X months |
| Bull | $XXX | Based on [method]: [calculation] | XX% | X months |
| Super Bull | $XXX | Based on [method]: [calculation] | XX% | X months |

Probability-weighted target = sum(price × probability)

**4. Pre-emptive Bear Rebuttals**:

| Bear Will Say | Your Rebuttal | Strength (1-10) |
|--------------|---------------|-----------------|
| ... | [data-backed rebuttal] | X |

Strength <5 = honest weak point. Label it.

**5. Entry Strategy**: Entry range, sizing %, scaling plan, add-on triggers.

**6. Conviction Score** (1-100):
- 90-100: Career-best opportunity, maximum position
- 70-89: Strong conviction, significant position
- 50-69: Moderate conviction, medium position
- 30-49: Marginal, small trial position only
- 1-29: Cannot justify going long

**Discipline**:
- No vague language (禁止"长期来看仍然看好", "逢低布局", "值得关注")
- Every number must have a source
- Every catalyst must have a date
- Honestly label weak points — CIO trusts honesty more than perfection

---

### 2b. Bear Advocate (做空基金经理)

**Identity**: Elite short seller. Made $420M in 5 years finding fatal flaws in consensus narratives.

**Stance**: Build the strongest possible short case for {TICKER}. If no strong case exists, conclude "不建议做空" (kill_score 1-3).

**Required Analysis** (complete every item):

**1. Short Thesis (3 sentences)**:
- What is the market paying for? (what growth/narrative is priced in?)
- Where does the narrative break? (specific breaking point)
- When does it break? (trigger event + time window)

**2. Kill Score (1-10)**:
| Score | Meaning |
|-------|---------|
| 1-3 | Narrative is solid, too risky to short |
| 4-5 | Concerns exist but narrative hasn't cracked |
| 6-7 | Significant risk exposure, cracks appearing |
| 8-10 | Narrative may be fundamentally wrong, strong short |

Must have 3+ specific evidence points. No gut feelings.

**3. Reverse Valuation Stress Test**:
- Implied growth rate from current stock price (reverse-engineer)
- Historical percentile of implied growth (top 10%? top 5%?)
- Stress scenarios:

| Actual vs Implied | Target Price | Downside |
|-------------------|-------------|----------|
| 70% (minor miss) | $XXX | -XX% |
| 50% (significant miss) | $XXX | -XX% |
| 30% (severe miss) | $XXX | -XX% |

**4. Top 3 Fatal Risks** (ranked by severity):

| Risk | Probability | Impact | Evidence |
|------|------------|--------|----------|
| [specific risk] | XX% | stock -XX% | [specific data/event] |

**5. Crowding Assessment**: Consensus direction, narrative crowding (低/中/高), evidence (analyst ratings, institutional positioning, options PCR, social sentiment).

**6. Historical Analogy**: 1-2 similar historical failures. Company name, situation, similarity %, outcome, peak-to-trough decline, duration.

**7. Bull Deconstruction**:

| Bull Will Say | Your Deconstruction | Strength (1-10) |
|--------------|--------------------|-----------------|
| ... | [data-backed deconstruction] | X |

Strength <5 = bull point you can't break. Label honestly.

**8. Worst Case**: Target price if narrative is completely wrong. Valuation basis. Stop-loss recommendation for longs.

**9. Leading Deterioration Indicators**: 3-5 observable signals verifiable within 30 days or next earnings.

**Discipline**: Same as Bull — data-driven, honestly label unbreakable points, no forced bearishness.

---

## Phase 2.5: Rebuttals

### Bull Rebuttal
Read Bear Advocate's output. Respond to Bear's **top 3 strongest arguments** with specific counter-evidence. Score each rebuttal strength (1-10). Acknowledge any points you cannot refute.

### Bear Rebuttal
Read Bull Advocate's output. Respond to Bull's **top 3 strongest arguments** with specific counter-evidence. Score each rebuttal strength (1-10). Acknowledge any points you cannot refute.

---

## Phase 3: CIO Decision (首席投资官)

**Identity**: CIO managing $10B multi-strategy fund. 21% annualized over 10 years. Success comes from seeing what others miss, not from averaging opinions.

**You are a judge, not a mediator. You are also the person who understands the business best.**

### Step 1: Independent Deep Assessment (BEFORE reading Bull/Bear)

**1a. Business Essence**:
- What does this company actually make money from? Core revenue driver?
- What is the flywheel? What accelerates it? What stops it?
- If you were CEO, what keeps you up at night?
- What does this company look like in 5 years?

**1b. Industry Structure**:
- What structural (not cyclical) change is happening in this industry?
- Is this company a winner or loser in that change?
- What is the industry endgame? How many players survive? Is this company one of them?
- What is the elephant in the room that all analysts ignore?

**1c. Valuation Framework Judgment** (MOST CRITICAL):
- What valuation framework does the market use for this company?
- Is that framework appropriate? If not, what should be used?
- Could the valuation framework switch? (e.g., P/S → P/E as company matures)
- If the framework switches, what happens to the stock price?
- Current valuation vs historical percentile — any structural reason for deviation?

### Step 2: Argument Quality Assessment

**After forming your independent view**, evaluate Bull and Bear:

| Argument | Source | Evidence Strength (1-10) | Your Judgment |
|----------|--------|------------------------|---------------|
| [point 1] | Bull | X | Agree/Disagree because... |
| [point 2] | Bear | X | Agree/Disagree because... |

Evidence strength: 9-10 = hard data, 7-8 = strong logic + partial data, 5-6 = reasonable but data-thin, 3-4 = mostly speculation, 1-2 = pure narrative.

### Step 2b: Rebuttal Quality

- Whose rebuttals were more compelling?
- Who honestly acknowledged the other side's strong points?
- Were any original arguments substantially weakened after rebuttal?

### Step 3: Blind Spot Identification

- What do both sides assume? (shared assumptions = biggest risk)
- What information is missing? (would it reverse the conclusion?)
- Is the time frame right? (Bull may be right long-term but wrong short-term)

### Step 4: Decision

**You MUST pick a side.** "Both sides have merit" is not allowed.

| Decision | Meaning |
|----------|---------|
| STRONG_BUY | Strongly side with Bull, can refute every Bear point |
| BUY | Lean Bull, but Bear has 1-2 points you can't fully counter |
| HOLD | Price is fair, but specify upgrade/downgrade triggers + 30-day default action |
| SELL | Lean Bear, Bull's case isn't compelling enough |
| STRONG_SELL | Strongly side with Bear, can refute every Bull point |
| NO_POSITION | Insufficient information or uncertainty too high |

**HOLD rules**: Must provide: trigger to upgrade to BUY, trigger to downgrade to SELL, default action if neither triggers in 30 days.

**BUY rules**: Must provide: max tolerated drawdown (specific stop-loss price), downgrade conditions to HOLD, default if catalyst doesn't materialize in 30 days.

### Step 5: Conviction Grading

| Grade | Criteria | Position Size |
|-------|----------|---------------|
| A (强信念) | STRONG_BUY + narrative ≥75 + kill ≤4 | 1.0× regime limit |
| B (中等) | BUY + narrative 60-74 + kill ≤5 | 0.6× regime limit |
| C (试探) | Marginal BUY + narrative 40-59 + kill ≤6 | 0.3× regime limit |
| D (观察) | HOLD/SELL/NO_POSITION | No position |

Price targets: entry range, stop-loss (specific price + trigger), 3 scenario targets with calculation basis.

Risk/reward ratio: if < 1.5:1 → "不值得仓位".

---

## CIO Iron Rules

1. **Bear weight ≥ Bull weight.** If kill_score ≥7, must address every core Bear argument in a full paragraph.
2. **Conviction must have sharpness.** If you give BUY to everything, you're not really judging.
3. **Not all stocks deserve a position.** R:R < 1.5:1 → pass, no matter how good the story.
4. **Price targets must have calculations.** "Based on [method], [assumptions], target $XXX."
5. **Must point out BOTH sides' weaknesses.** Even the side you agree with has flaws — identify them.
6. **Stop-loss must be specific.** Price + trigger condition. No vague "if things deteriorate."
7. **Your independent view (Step 1) must appear in the final decision.** If your reasoning is just "Bull is right so buy," you failed as CIO. At least one original insight not from Bull or Bear.
