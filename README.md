# AI-Powered Consolidated Variance Analysis

A Python agent that automates the quarterly close consolidation workflow for public companies. It reads department-level financial submissions, consolidates P&L and Balance Sheet data across operating segments and corporate functions, identifies material variances, and generates SEC-quality variance commentary using the Claude API — including catching errors in department explanations.

## What It Does

In a public company quarterly close, every department submits their financials with explanations for material variances. The corporate accounting team then consolidates everything, writes the variance analysis, and catches mistakes before it goes to the CFO and external auditors.

This tool automates that entire workflow:

1. **Reads** department submission workbooks (3 operating segments + 4 corporate functions)
2. **Consolidates** across all units into a single P&L (quarterly + YTD) and Balance Sheet
3. **Identifies** material variances using configurable thresholds (default: >5% or >$500K)
4. **Generates** AI-written variance commentary with:
   - Dollar amounts on every driver, ordered by magnitude
   - Department attribution ranked by contribution size
   - 80%+ coverage target per line item
5. **Catches errors** in department explanations:
   - Directional mistakes (says "favorable" when expense increased)
   - Math that doesn't tie (driver dollars don't add to total variance)
   - Contradictions (claims balance decreased but it actually increased)
   - Inflated or fabricated dollar amounts

## Output

The agent generates a consolidated Excel report with 4 tabs:

| Tab | Description |
|-----|-------------|
| **Consolidated Qtr PL** | Quarterly P&L with QoQ and YoY variances, AI explanations, and AI-generated follow-up flags |
| **Consolidated YTD PL** | Year-to-date P&L with H1 2025 vs H1 2024 analysis |
| **Consolidated BS** | Balance Sheet with Jun 30, 2025 vs Dec 31, 2024 analysis |
| **Dept Summary** | Department contribution counts |

- **Explanation column**: Clean variance narrative only — no caveats, no flags
- **Follow-up column**: AI-generated only — populated when the AI catches errors or insufficient coverage

## Architecture

```
dept_submissions_input.xlsx          consolidated_narrator.py          consolidated_variance_report.xlsx
┌──────────────────────────┐        ┌─────────────────────┐          ┌────────────────────────────────┐
│ Enterprise Solutions     │        │                     │          │ Consolidated Qtr PL            │
│   Qtr PL / YTD PL / BS  │        │  1. Parse input     │          │   QoQ + YoY explanations       │
│ Cloud & Service Provider │───────>│  2. Consolidate     │────────> │ Consolidated YTD PL            │
│   Qtr PL / YTD PL / BS  │        │  3. Claude API      │          │   H1 vs H1 explanations        │
│ Security & AI            │        │  4. Build output    │          │ Consolidated BS                │
│   Qtr PL / YTD PL / BS  │        │                     │          │   Period-over-period            │
│ Legal / HR / Corp Finance│        └─────────────────────┘          │ Dept Summary                   │
│ IT & Operations          │                  │                      └────────────────────────────────┘
└──────────────────────────┘                  │
                                     Claude API (Sonnet)
                                     - Synthesizes dept explanations
                                     - Validates math & direction
                                     - Catches errors
                                     - Generates follow-up flags
```

## Usage

```bash
# Install dependencies
pip install -r requirements.txt

# Run with API key file (put key in API_Key.txt in same folder)
python consolidated_narrator.py dept_submissions_input.xlsx

# Or pass key directly
python consolidated_narrator.py dept_submissions_input.xlsx --api-key sk-ant-your-key

# Custom output and thresholds
python consolidated_narrator.py dept_submissions_input.xlsx -o Q2_report.xlsx --materiality-pct 0.10 --materiality-abs 1000000

# Skip AI (consolidation only, dept explanations passed through)
python consolidated_narrator.py dept_submissions_input.xlsx --no-ai
```

### API Key Setup

The agent looks for your Anthropic API key in this order:
1. `--api-key` command line flag
2. `ANTHROPIC_API_KEY` environment variable
3. `API_Key.txt` file in the same folder as the script
4. Interactive prompt (paste key, not stored)

## Department Submission Format

Each department submits 3 tabs in the input workbook:

- **Qtr PL**: Quarterly actuals (Q2 2025, Q1 2025, Q2 2024) with QoQ and YoY explanations for material variances
- **YTD PL**: Year-to-date actuals (H1 2025, H1 2024) with YTD explanations
- **BS**: Balance Sheet (Jun 30, 2025 vs Dec 31, 2024) with period-over-period explanations

Departments submit **final** numbers and explanations — no follow-up flags in the input. The AI generates follow-up flags on the consolidated output only when it finds issues.

## Materiality Thresholds

Default: variance is material if **>5% OR >$500K** (absolute). Both thresholds are configurable via command line and displayed in the output headers.

## Tech Stack

- Python 3.12+
- [Anthropic Claude API](https://docs.anthropic.com/) (Sonnet) for variance commentary
- openpyxl for Excel generation
- pandas for data processing

## Business Segments Covered

**Operating Segments:**
- Enterprise Solutions (campus networking, SD-WAN, Mist AI, wireless)
- Cloud & Service Provider (carrier routing, 5G, data center switching, Apstra/Paragon)
- Security & AI (SRX firewalls, threat intelligence, managed security, AI/ML)

**Corporate Functions:**
- Legal (outside counsel, litigation, IP, compliance)
- Human Resources (compensation, benefits, recruitment, severance)
- Corporate Finance (treasury, tax, audit, debt, equity, intercompany)
- IT & Operations (cloud infrastructure, cybersecurity, facilities)

## Status

Active development. Core consolidation and AI commentary engine functional. Iterating on output formatting and prompt engineering.

## Author

**Joel Stell, CPA MBA**
