# Setup Guide

## Environment Variables

Create a `.env` file or export these variables in your shell:

```bash
# Required for chain-screener, stock-screener, equity-research
export FMP_API_KEY="your-key-here"
# Register at: https://financialmodelingprep.com/developer/docs/

# Required for market-sizing (LLM analysis)
export GEMINI_API_KEY="your-key-here"
# Register at: https://ai.google.dev/

# Required for equity-research, market-sizing (alternative LLM)
export ANTHROPIC_API_KEY="your-key-here"
# Register at: https://console.anthropic.com/

# Optional: China/HK stock data for chain-screener
export TUSHARE_TOKEN="your-token-here"
# Register at: https://tushare.pro/

# Optional: path to your quant trading system root
export QUANT_ROOT="/path/to/your/quant-trading"
```

## Python Dependencies

```bash
pip install python-pptx openpyxl requests
```

## Usage with Claude Code

### Install as Skills

```bash
# Copy templates to Claude Code skills directory
cp -r templates/cathay-ppt ~/.claude/skills/cathay-ppt-template
cp -r templates/cathay-excel ~/.claude/skills/cathay-excel-template
cp -r skills/equity-research ~/.claude/skills/equity-research

# For OpenClaw agent workspace skills
cp -r skills/market-sizing ~/.openclaw/workspace/skills/
cp -r skills/chain-screener ~/.openclaw/workspace/skills/
cp -r skills/stock-screener ~/.openclaw/workspace/skills/
cp -r skills/stock-compare ~/.openclaw/workspace/skills/
```

### Verify Installation

In Claude Code, these skills should appear in your skill list:
- `/cathay-ppt-template` — Generate Cathay-branded decks
- `/cathay-excel-template` — Build PE financial models
- `/equity-research` — Run equity analysis with MoA debate

## API Registration Guide

### FMP (Financial Modeling Prep)
1. Go to https://financialmodelingprep.com/developer/docs/
2. Sign up for a free account (250 calls/day)
3. Copy your API key from the dashboard
4. For production use, consider the Starter plan ($14/mo, 750 calls/day)

### Google Gemini
1. Go to https://ai.google.dev/
2. Create a project in Google Cloud Console
3. Enable the Generative AI API
4. Generate an API key

### Anthropic Claude
1. Go to https://console.anthropic.com/
2. Create an account and add billing
3. Generate an API key from Settings → API Keys

### Tushare Pro (Optional)
1. Go to https://tushare.pro/
2. Register with phone number
3. Get token from user center
4. Note: Some endpoints require accumulated credits (积分)
