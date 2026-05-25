"""Apply current-events reframing and 5 labeled numerical choices to chapter Excel files."""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

BASE = Path(__file__).parent
HEADER_FILL = PatternFill("solid", fgColor="2F5496")
HEADER_FONT = Font(name="Arial", size=11, bold=True, color="FFFFFF")
BODY_FONT = Font(name="Arial", size=10)
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

CURRENT_EVENTS = {
    "Chapter_Introduction_100_Questions.xlsx": {
        81: (
            "Major US banks left the Net Zero Banking Alliance in 2025; BlackRock's exit paused the Net Zero Asset Managers initiative. "
            "Write a short commentary: what do these withdrawals imply about (i) voluntary climate coalitions versus binding fiduciary duties, "
            "(ii) the gap between financed-emissions rhetoric and balance-sheet reality, and (iii) reputational risk when public commitments outpace measurable decarbonisation?"
        ),
        82: (
            "The EU deforestation regulation was postponed with exemptions, and many member states had not transposed the CSRD by late 2024. "
            "Comment on whether this pattern reflects regulatory backlash, implementation fatigue, or a deliberate competitiveness recalibration — "
            "and how it affects investor confidence in the EU sustainable-finance rulebook (SFDR/CSRD/Taxonomy) described in the chapter."
        ),
        83: (
            "Hundreds of firms withdrew SBTi net-zero targets, citing Scope 3 data and accountability limits. "
            "Explain why Scope 3 is structurally harder than Scope 1–2 for target-setting and disclosure, "
            "and whether dropping targets strengthens or weakens the credibility of corporate climate claims."
        ),
        84: (
            "Microsoft reported rising emissions (~30% vs. 2020) while expanding AI data centres. "
            "Discuss the tension between absolute decoupling, digital demand growth, and corporate net-zero pledges — "
            "does this case support or undermine the chapter's argument that efficiency gains rarely offset scale effects?"
        ),
        85: (
            "At the ICJ (December 2024), major emitters argued the Paris Agreement sets political expectations, not enforceable legal duties. "
            "Comment on how this aligns with the chapter's view of weak global governance, free-riding, and the limits of voluntary climate architecture."
        ),
        86: (
            "Walmart rolled back DEI programmes after political pressure in 2024. "
            "Interpret this through the chapter's instrumental versus intrinsic sustainability distinction and the idea of cosmetic or reversible 'S' policies when social risk is not priced."
        ),
        87: (
            "European farmer protests (2023–2024) targeted sustainability rules (CAP green requirements, pesticide/nitrate constraints). "
            "Comment on the political economy of imposing transition costs on dispersed producers while urban consumers and investors demand greener supply chains."
        ),
        88: (
            "COP29's Loss and Damage Fund was set near $300B/year, far below developing countries' ~$1.3T request. "
            "Relate the outcome to the tragedy of the commons, North–South bargaining power, and why ex-post compensation does not replace ex-ante mitigation incentives."
        ),
        89: (
            "Amazon was removed from SBTi's validated list for lacking a credible reduction pathway. "
            "Discuss the gap between Business Roundtable-style stakeholder rhetoric and verifiable emissions trajectories — what does delisting signal for 'commitment without delivery'?"
        ),
        90: (
            "The Draghi Report argued EU sustainability regulation hurts competitiveness. "
            "Comment on the trade-off between compliance costs (CSRD, due diligence) and long-run competitiveness, linking to the chapter's ESG backlash debate and Friedman-style shareholder primacy."
        ),
    },
    "Chapter_ESG_Scores_100_Questions.xlsx": {
        81: (
            "S&P Global discontinued ESG credit indicators in 2023. "
            "Comment on what this signals about ESG data quality, the ~0.54 cross-provider correlation, and whether ESG metrics are ready to drive credit spreads as traditional ratings do."
        ),
        82: (
            "The EU adopted the ESGR regulation (2024), the first regime for ESG rating providers. "
            "Explain why opacity, scope/measurement divergence, and issuer-pay conflicts of interest made provider regulation necessary — reference the chapter's provider business models."
        ),
        83: (
            "Italian prosecutors documented labour exploitation in LVMH's supply chain (2024). "
            "Comment on how controversy-based 'S' risks can be invisible in headline ESG scores when measurement relies on sparse supply-chain data and best-in-class peer ranking."
        ),
        84: (
            "Nestlé marketed sweeter infant formula in lower-income markets than in Europe. "
            "Contrast how a financial-materiality lens versus an impact-materiality lens would treat this practice — which stakeholders' welfare enters each definition?"
        ),
        85: (
            "The Net Zero Banking Alliance unravelled as major banks exited (2024–2025). "
            "Discuss stated-versus-real financed emissions and whether alliance membership should have lifted ESG scores if decarbonisation pathways were not verified."
        ),
        86: (
            "VW sold Xinjiang facilities after a discredited audit (2024). "
            "Comment on score manipulation, controversy lag, and whether divestiture resolves governance concerns or mainly manages reputational exposure."
        ),
        87: (
            "MSCI kept high ratings on banks with continued Russian exposure after 2022 (e.g. Société Générale, Metro per Ahmed et al.). "
            "What does this imply about ratings' responsiveness to geopolitical/controversy shocks versus slow-moving ESG inputs?"
        ),
        88: (
            "Elon Musk called ESG a 'scam' after Tesla's S&P 500 ESG Index removal while tobacco names scored highly. "
            "Evaluate the criticism using the chapter's weighting/scope disagreements and sector-neutral best-in-class logic — which parts are fair, which are misplaced?"
        ),
        89: (
            "The EU CSRD (2024) expands double materiality reporting to ~50,000 firms. "
            "Explain how reporting impact materiality alongside financial materiality could narrow the 'Big Decorrelation' between providers — and what limits remain."
        ),
        90: (
            "Amundi pressed LVMH for stricter oversight after the supply-chain scandal. "
            "Assess whether stewardship/engagement can substitute for flawed third-party ESG scores, given the chapter's limits on investor influence and data latency."
        ),
    },
}

NUMERICAL_CHOICES = {
    "Chapter_Introduction_100_Questions.xlsx": {
        71: [
            "NPV = +500M",
            "NPV = +632.3M",
            "NPV = +1,400M",
            "NPV = +1,032.3M",
            "NPV = −310M",
        ],
        72: [
            "NPV = −1,088.2M",
            "NPV = +632.3M",
            "NPV = −550M",
            "NPV = −688.2M",
            "NPV = −1,388.2M",
        ],
        73: [
            "NPV = +215.5M",
            "NPV = +815.5M",
            "NPV = +632.3M",
            "NPV = +1,150M",
            "NPV = +115.5M",
        ],
        74: [
            "Change in PV = 339.6M",
            "Change in PV = 231.6M",
            "Change in PV = 280.4M",
            "Change in PV = 419.6M",
            "Change in PV = 159.6M",
        ],
        75: [
            "Reduction = 27.0%",
            "Reduction = 22.0%",
            "Reduction = 25.0%",
            "Reduction = 30.0%",
            "Reduction = 24.5%",
        ],
        76: [
            "Carbon burden = $5,236M",
            "Carbon burden = $4,636M",
            "Carbon burden = $5,636M",
            "Carbon burden = $6,636M",
            "Carbon burden = $5,836M",
        ],
        77: [
            "Additional revenue = $735M",
            "Additional revenue = $750M",
            "Additional revenue = $650M",
            "Additional revenue = $685M",
            "Additional revenue = $725M",
        ],
        78: [
            "Increase = 300 TWh",
            "Increase = 240 TWh",
            "Increase = 360 TWh",
            "Increase = 180 TWh",
            "Increase = 320 TWh",
        ],
        79: [
            "NPV = +441M",
            "NPV = −500M",
            "NPV = −380M",
            "NPV = −620M",
            "NPV = −154M",
        ],
        80: [
            "Number of firms = 3,080",
            "Number of firms = 2,800",
            "Number of firms = 3,200",
            "Number of firms = 2,500",
            "Number of firms = 3,500",
        ],
    },
    "Chapter_ESG_Scores_100_Questions.xlsx": {
        71: [
            "Overlap = 63%",
            "Overlap = 54%",
            "Overlap = 80%",
            "Overlap = 45%",
            "Overlap = 72%",
        ],
        72: [
            "Boolean indicators = 112",
            "Boolean indicators = 74",
            "Boolean indicators = 93",
            "Boolean indicators = 130",
            "Boolean indicators = 100",
        ],
        73: [
            "Remaining divergence = 62%",
            "Remaining divergence = 38%",
            "Remaining divergence = 56%",
            "Remaining divergence = 44%",
            "Remaining divergence = 50%",
        ],
        74: [
            "Percentile disagreement = 70 pp",
            "Percentile disagreement = 54 pp",
            "Percentile disagreement = 30 pp",
            "Percentile disagreement = 85 pp",
            "Percentile disagreement = 60 pp",
        ],
        75: [
            "Share used in score = 21.4%",
            "Share used in score = 25.0%",
            "Share used in score = 33.3%",
            "Share used in score = 15.0%",
            "Share used in score = 40.0%",
        ],
        76: [
            "Share to expansion = 49.4%",
            "Share to expansion = 45.0%",
            "Share to expansion = 55.0%",
            "Share to expansion = 40.0%",
            "Share to expansion = 52.0%",
        ],
        77: [
            "Price/assembly ratio = 49.1×",
            "Price/assembly ratio = 35×",
            "Price/assembly ratio = 26×",
            "Price/assembly ratio = 55×",
            "Price/assembly ratio = 42×",
        ],
        78: [
            "Companies rated A = 100",
            "Companies rated A = 150",
            "Companies rated A = 75",
            "Companies rated A = 125",
            "Companies rated A = 200",
        ],
        79: [
            "Total compliance cost = €50 billion",
            "Total compliance cost = €25 billion",
            "Total compliance cost = €10 billion",
            "Total compliance cost = €75 billion",
            "Total compliance cost = €40 billion",
        ],
        80: [
            "Expected percentile = 75th",
            "Expected percentile = 90th",
            "Expected percentile = 65th",
            "Expected percentile = 80th",
            "Expected percentile = 70th",
        ],
    },
}

CORRECT_INDEX = {
    "Chapter_Introduction_100_Questions.xlsx": {71: 2, 72: 1, 73: 1, 74: 1, 75: 1, 76: 3, 77: 1, 78: 1, 79: 2, 80: 1},
    "Chapter_ESG_Scores_100_Questions.xlsx": {71: 1, 72: 1, 73: 1, 74: 1, 75: 1, 76: 1, 77: 1, 78: 1, 79: 1, 80: 1},
}


def style_header_row(ws, ncol):
    for col in range(1, ncol + 1):
        c = ws.cell(1, col)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER


def update_file(filename):
    path = BASE / filename
    wb = openpyxl.load_workbook(path)
    ws = wb["All Questions"]
    for qid, text in CURRENT_EVENTS.get(filename, {}).items():
        for row in ws.iter_rows(min_row=2):
            if row[0].value == qid:
                row[2].value = text
                break

    if "Numerical Choices" in wb.sheetnames:
        del wb["Numerical Choices"]
    ws2 = wb.create_sheet("Numerical Choices")
    headers = ["#", "Question", "Model Answer"] + [f"Choice {i}" for i in range(1, 6)] + ["Correct Choice #"]
    ws2.append(headers)
    style_header_row(ws2, len(headers))

    ws_q = wb["All Questions"]
    qmap = {}
    for row in ws_q.iter_rows(min_row=2, values_only=True):
        if row and row[0] and (row[1] or "").strip() == "Numerical":
            qmap[row[0]] = (row[2], row[3])

    choices_data = NUMERICAL_CHOICES[filename]
    correct = CORRECT_INDEX[filename]
    for qid in sorted(choices_data.keys()):
        q_text, ans = qmap.get(qid, ("", ""))
        row = [qid, q_text, ans] + choices_data[qid] + [correct[qid]]
        ws2.append(row)
        for col in range(1, len(row) + 1):
            c = ws2.cell(ws2.max_row, col)
            c.font = BODY_FONT
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.border = BORDER
        ws2.cell(ws2.max_row, 9).font = Font(name="Arial", size=10, bold=True, color="2F5496")

    ws2.freeze_panes = "A2"
    try:
        wb.save(path)
        print(f"Updated {filename}")
    except PermissionError:
        alt = path.with_name(path.stem + "_updated.xlsx")
        wb.save(alt)
        print(f"Could not overwrite {filename} (file open?). Saved {alt.name}")


if __name__ == "__main__":
    for f in CURRENT_EVENTS:
        update_file(f)
