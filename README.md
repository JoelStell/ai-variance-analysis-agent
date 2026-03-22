# AI Variance Analysis Agent

A Python agent that automates the quarterly close consolidation workflow for public companies. It reads department-level financial submissions across operating segments and corporate functions, consolidates P&L and Balance Sheet data, identifies material variances, and generates SEC-quality variance commentary using the Claude API — including catching errors in department explanations.

## What It Does

In a public company quarterly close, every department submits their financials with explanations for material variances. The corporate accounting team then consolidates everything, writes the variance analysis, and catches mistakes before it goes to the CFO and external auditors.

This tool automates that entire workflow:

1. **Reads** department submission workbooks (3 operating segments + 4 corporate functions)
2. **Consolidates** across all units into Quarterly P&L, Year-to-Date P&L, and Balance Sheet
3. **Identifies** material variances using configurable thresholds (default: >5% or >$500K)
4. **Generates** AI-written variance commentary with:
   - Dollar amounts on every driver, ordered by magnitude
   - Department attribution ranked by contribution size
   - Coverage percentage showing what portion of each variance is explained
5. **Catches errors** in department explanations:
   - Directional mistakes (says "favorable" when expense increased)
   - Math that doesn't reconcile (driver dollars don't add to total variance)
   - Contradictions between explanation and actual data
   - Insufficient coverage (below 80% threshold)
6. **Generates follow-up actions** organized by department, ready to email

## Output

The agent generates a consolidated Excel report with 4 tabs:

| Tab | Description |
|-----|-------------|
| **Consolidated Qtr PL** | Quarter-over-quarter and year-over-year analysis side by side with self-contained comparison blocks, AI explanations, coverage %, and follow-up flags |
| **Consolidated YTD PL** | Six months ended June 30, 2025 vs June 30, 2024 with YTD explanations |
| **Consolidated BS** | Balance Sheet as of June 30, 2025 vs December 31, 2024 |
| **Follow-Up Actions** | All AI-flagged issues grouped by department with specific action required and status tracking |

**Design principles:**
- Explanation column contains clean variance narrative only — no flags or caveats
- Follow-up column is AI-generated only — populated when the AI catches errors or insufficient coverage
- Coverage % in its own column (green ≥80%, orange <80%) — not embedded in explanation text
- Quarterly P&L repeats the current period column before YoY section so the reader never scans back across the page
- Follow-up flags identify which specific department to contact

## Architecture

```
dept_submissions_input.xlsx            consolidated_narrator.py         consolidated_variance_report.xlsx
┌────────────────────────────┐        ┌──────────────────────┐        ┌──────────────────────────────┐
│ 7 Departments × 3 tabs each│        │                      │        │ Consolidated Qtr PL          │
│                            │        │  1. Parse input       │        │   QoQ block | YoY block      │
│ Enterprise Solutions       │        │  2. Consolidate       │        │   Expl + Cov% + Flags        │
│   Qtr PL / YTD PL / BS    │───────>│  3. Claude API        │──────> │ Consolidated YTD PL          │
│ Cloud & Service Provider   │        │  4. Validate & catch  │        │ Consolidated BS              │
│ Security & AI              │        │  5. Build output      │        │ Follow-Up Actions            │
│ Legal / HR / Corp Finance  │        │                       │        │   Grouped by department      │
│ IT & Operations            │        └──────────────────────┘        └──────────────────────────────┘
└────────────────────────────┘                  │
                                       Claude API (Sonnet)
                                       Synthesizes explanations
                                       Validates math & direction
                                       Catches errors
                                       Assigns follow-up by dept
```

## Usage

```bash
# Install dependencies
pip install -r requirements.txt

# Run with API key file (put key in API_Key.txt in same folder)
python consolidated_narrator.py dept_submissions_input.xlsx

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
4. Interactive prompt

## Department Submission Format

Each department submits 3 tabs in the input workbook:

- **Qtr PL**: Quarterly actuals with QoQ and YoY explanations for material variances
- **YTD PL**: Year-to-date actuals (Six Months Ended June 30) with YTD explanations
- **BS**: Balance Sheet with period-over-period explanations

Departments submit **final** numbers and explanations. There are no follow-up flags in the input — the AI generates follow-up flags on the consolidated output only when it finds issues.

## Materiality Thresholds

Default: variance is material if **>5% OR >$500K** (absolute). Both thresholds are configurable via command line and displayed in the output headers.

Explanations are required to cover at least **80%** of the dollar variance. The AI tracks coverage percentage independently and flags items below the threshold.

## Tech Stack

- Python 3.12+
- [Anthropic Claude API](https://docs.anthropic.com/) (Sonnet) for variance commentary and error detection
- openpyxl for Excel generation
- pandas for data processing

## Business Units Covered

**Operating Segments:**
- Enterprise Solutions (campus networking, SD-WAN, Mist AI, wireless)
- Cloud & Service Provider (carrier routing, 5G, data center switching, Apstra/Paragon)
- Security & AI (SRX firewalls, threat intelligence, managed security, AI/ML)

**Corporate Functions:**
- Legal (outside counsel, litigation, IP, compliance)
- Human Resources (compensation, benefits, recruitment, severance)
- Corporate Finance (treasury, tax, audit, debt, equity, intercompany)
- IT & Operations (cloud infrastructure, cybersecurity, facilities)

## Author

**Joel Stell, CPA MBA**
