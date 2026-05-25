"""Rewrite Numerical Choices with step-by-step calculations (not results only)."""
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

BASE = Path(__file__).parent
THIN = Side(style="thin")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BODY_FONT = Font(name="Arial", size=10)

def ch_intro_en():
    return {
        71: (["NPV = −400 + 90×10 = −400 + 900 = +500M (10-yr annuity, not 20)",
               "NPV = −400 + 90×11.47 = −400 + 1,032.3 = +632.3M",
               "NPV = −400 + 90×20 = −400 + 1,800 = +1,400M (undiscounted)",
               "NPV = 90×11.47 = +1,032.3M (omits initial investment)",
               "NPV = −400 + 90 = −310M (single cash flow)"], 2),
        72: (["Carbon cost = 1.5 Mt × €100/t = €150M/yr → net CF = 90−150 = −€60M/yr → NPV = −400+(−60)×11.47 = −400−688.2 = −€1,088.2M",
               "NPV = −400 + 90×11.47 = +632.3M (ignores carbon)",
               "NPV = −400 − 150 = −550M (one-off carbon, not annual)",
               "Net CF = −60M/yr only; PV = −60×11.47 = −688.2M (no I₀)",
               "NPV = −400 + (90−150−150)×11.47 (double-counts carbon)"], 1),
        73: (["NPV = −600 + 70×11.65 = −600 + 815.5 = +215.5M",
               "NPV = 70×11.65 = +815.5M (no capex)",
               "NPV = −600 + 70×25 = +1,150M (undiscounted)",
               "NPV = +632.3M (gas plant, wrong project)",
               "NPV = −600 + 70×10 = +100M (wrong annuity factor)"], 1),
        74: (["ΔPV = 1,280.4 − 940.8 = 339.6M", "ΔPV = 940.8 − 1,280.4 = −339.6M (wrong sign)",
               "Rate gap only: 12% − 8% = 4 pp", "ΔPV = 120×(10.67−7.84) = 339.6M", "ΔPV = 280.4M"], 1),
        75: (["Reduction = (4.92−3.59)/4.92 = 1.33/4.92 = 27.0%",
               "Absolute drop = 4.92−3.59 = 1.33 Gt (not %)", "Remaining share = 3.59/4.92 = 73%",
               "Reduction = 25.0%", "Reduction = 30.0%"], 1),
        76: (["Annual cost = 5 Mt × $80/t = $400M/yr", "Carbon burden = 400×14.09 = $5,636M",
               "Burden / market cap = 5,636/4,000 = 141%", "Burden = 400×25 = $10,000M (undiscounted)",
               "Burden per ton only"], 2),
        77: (["Added revenue = (750−15)×10,000×100 = $735M/yr", "New revenue = $750M/yr",
               "Old revenue = $15M/yr", "Per patient = $73,500/yr", "Price ratio = 50×"], 1),
        78: (["2025: 4.5%×4,000 = 180 TWh; 2035: >10%×4,000 > 400 TWh → increase ≥ 300 TWh",
               "Increase = 240 TWh", "Increase = 360 TWh", "Increase = 180 TWh only",
               "Share points: 5.5 pp only"], 1),
        79: (["Carbon = 2 Mt×€60 = €120M/yr → net CF = 0 → NPV = −500+0×7.84 = −500M",
               "NPV = −500 + 120×7.84 = +441M (no carbon)", "NPV = −500 + 120×10.67 = +780M",
               "NPV = −500 − 120 = −620M", "NPV = −154M (solar reference)"], 1),
        80: (["Count = 77%×4,000 = 3,080 firms", "Share = 77% only", "Below cap = 920 firms",
               "Invalid aggregate", "Count = 3,500"], 1),
    }

def ch_intro_fr():
    return {
        71: (["VAN = −400 + 90×10 = −400 + 900 = +500 M€ (annuité 10 ans, pas 20)",
               "VAN = −400 + 90×11,47 = −400 + 1 032,3 = +632,3 M€",
               "VAN = −400 + 90×20 = −400 + 1 800 = +1 400 M€ (sans actualisation)",
               "VAN = 90×11,47 = +1 032,3 M€ (oublie l'investissement)",
               "VAN = −400 + 90 = −310 M€ (un seul flux)"], 2),
        72: (["Coût carbone = 1,5 Mt × 100 €/t = 150 M€/an → CF net = 90−150 = −60 M€/an → VAN = −400+(−60)×11,47 = −400−688,2 = −1 088,2 M€",
               "VAN = −400 + 90×11,47 = +632,3 M€ (ignore le carbone)",
               "VAN = −400 − 150 = −550 M€ (carbone ponctuel)",
               "CF net = −60 M€/an ; PV = −60×11,47 = −688,2 M€ (sans I₀)",
               "VAN = −400 + (90−150−150)×11,47 (double comptage)"], 1),
        73: (["VAN = −600 + 70×11,65 = −600 + 815,5 = +215,5 M€",
               "VAN = 70×11,65 = +815,5 M€ (sans investissement)",
               "VAN = −600 + 70×25 = +1 150 M€ (non actualisé)",
               "VAN = +632,3 M€ (centrale gaz)",
               "VAN = −600 + 70×10 = +100 M€"], 1),
        74: (["ΔVA = 1 280,4 − 940,8 = 339,6 M€", "ΔVA = 940,8 − 1 280,4 = −339,6 M€",
               "Écart de taux : 4 points", "ΔVA = 120×(10,67−7,84) = 339,6 M€", "ΔVA = 280,4 M€"], 1),
        75: (["Réduction = (4,92−3,59)/4,92 = 27,0 %", "Écart absolu = 1,33 Gt",
               "Part restante = 73 %", "Réduction = 25,0 %", "Réduction = 30,0 %"], 1),
        76: (["Coût annuel = 5 Mt × 80 $/t = 400 M$/an", "Fardeau = 400×14,09 = 5 636 M$",
               "Fardeau / cap. boursière = 5 636/4 000 = 141 %", "Fardeau = 400×25 M$", "Par tonne seulement"], 2),
        77: (["Revenu suppl. = (750−15)×10 000×100 = 735 M$/an", "Revenu nouveau = 750 M$/an",
               "Revenu ancien = 15 M$/an", "Par patient = 73 500 $/an", "Ratio prix = 50×"], 1),
        78: (["2025 : 4,5%×4 000 = 180 TWh ; 2035 : >10%×4 000 → hausse ≥ 300 TWh",
               "Hausse = 240 TWh", "Hausse = 360 TWh", "Hausse = 180 TWh", "Écart 5,5 points seulement"], 1),
        79: (["Carbone = 2 Mt×60 €/t = 120 M€/an → CF net = 0 → VAN = −500 M€",
               "VAN = −500 + 120×7,84 = +441 M€", "VAN = −500 + 120×10,67 = +780 M€",
               "VAN = −500 − 120 = −620 M€", "VAN = −154 M€"], 1),
        80: (["Nombre = 77%×4 000 = 3 080", "Part = 77 %", "Sous le plafond = 920",
               "Agrégat invalide", "Nombre = 3 500"], 1),
    }

BONDS = {
    43: (["Penalty = 0.20%×€500M = €1M/yr ; Years 5–6 → 2×€1M = €2M", "€1M (1 year)", "€3M", "If obs. Y1: 5×€1M = €5M", "€2.5M"], "D"),
    44: (["Saving = 0.15%×€600M = €0.9M/yr ; PV factor sum ≈ 6.463 → NPV ≈ €5.82M", "€7.2M undisc.", "€4.95M", "€5.20M", "€5.82M"], "E"),
    45: (["PV savings Y1–6 ≈ €4.57M ; PV cost Y7–8 ≈ €0.83M → net NPV ≈ +€3.74M", "+€5.82M (met target)", "+€2.90M", "+€3.74M", "+€4.57M only"], "D"),
    46: (["Need PV penalty = PV savings → step-up ≈ 70 bp vs conventional", "25 bp enough", "~70 bp", "45 bp", "95 bp"], "C"),
    47: (["Y1–4 coupons €64M + Y5–7 €48M + fees €0.75M ≈ €112.75M", "€112M", "€113.5M", "€112.75M", "€115M no call", "€110.25M"], "C"),
    48: (["Surcharge = 0.50%×€775M = €3.875M/yr × 4 yrs = €15.5M", "€11.625M", "€15.5M", "€3.875M (1 yr)", "€19.375M"], "B"),
    49: (["Save 0.10%×€10B×8 = €80M ; penalty 0.25%×€10B×3 = €75M → net +€5M", "net −€5M", "net +€25M", "net +€5M", "obs. Y2: net −€70M", "+€80M save only"], "C"),
    50: (["Greenium 5bp: 0.05%×€2B = €1M/yr × 6 = €6M", "€12M", "€10M", "€3M", "€6M", "€1M/yr only"], "D"),
}

PERF = {
    41: (["Fund Sharpe = 12.2/18.5 = 0.659 ; Bench = 10.8/14 = 0.771", "0.771", "0.659", "0.743", "0.700"], "C"),
    42: (["α = 6.5−0.78×8.2 = +0.10%", "α = −1.60%", "IR = −0.29", "α = +0.60%", "α = +0.10%"], "D"),
    43: (["Sharpe Q = 8.8/13.2 = 0.667 ; Sharpe P = 16.5/28 = 0.589", "0.750", "0.540", "0.589", "0.667", "0.610"], "D"),
    44: (["α = 9.8−0.98×10.1 = −0.10% ; Sharpe ETF = 9.8/14.2 = 0.690", "−0.20%", "+0.10%", "+0.30%", "−0.10%", "−0.50%"], "D"),
    45: (["t(α)=1.29 < 1.96 → not significant ; net α = 2.16−1.90 = +0.26%", "−1.90%", "+0.18%/mo", "+2.16% headline", "t=1.29", "+0.26% net"], "E"),
    46: (["Sharpe X=9.2/16=0.575 ; Sharpe Y=9.2/22=0.418 ; αX=+0.28% ; αY=−2.28%", "0.475", "0.418", "0.650", "0.575", "0.500"], "D"),
    47: (["α = 5.5−6.5 = −1.0% ; Sharpe fund = 5.5/10.5 = 0.524 ; bench = 10/14 = 0.714", "+0.5%", "−2.5%", "−1.0%", "−0.5%", "+1.0%"], "C"),
    48: (["αA = 21−17.25 = +3.75% ; αB = 1−6.05 = −5.05% ; combined ≈ −0.62%", "−2.00%", "−0.62%", "+1.50%", "+3.75%", "−5.05%"], "D"),
    49: (["Sharpe fund = 9.5/13.8 = 0.688 ; bench = 10/14 = 0.714", "0.714", "0.688", "0.650", "0.730", "0.670"], "B"),
    50: (["Sharpe M = 4/20 = 0.200 ; αM = −3.65% ; Sharpe N = 8.5/14.5 = 0.586", "0.250", "0.450", "0.350", "0.586", "0.200"], "E"),
}

CONFIG = {
    "Chapter_Introduction_100_Questions.xlsx": ("12345", ch_intro_en()),
    "Chapter_Introduction_100_Questions_FR.xlsx": ("12345", ch_intro_fr()),
    "Chapter_FirmBonds_70_Questions.xlsx": ("ABCDE", BONDS),
    "Chapter_FirmBonds_70_Questions_FR.xlsx": ("ABCDE", BONDS),
    "Chapter_SustPerformance_70_Questions.xlsx": ("ABCDE", PERF),
    "Chapter_SustPerformance_70_Questions_FR.xlsx": ("ABCDE", PERF),
}


def rebuild_sheet(wb_path: Path, fmt: str, choices_map: dict):
    wb = openpyxl.load_workbook(wb_path)
    ws_q = wb["All Questions"]
    q_text, q_ans = {}, {}
    for row in ws_q.iter_rows(min_row=2, values_only=True):
        if row and row[0] is not None:
            q_text[row[0]] = row[2]
            q_ans[row[0]] = row[3]
    if "Numerical Choices" in wb.sheetnames:
        del wb["Numerical Choices"]
    ws2 = wb.create_sheet("Numerical Choices")
    if fmt == "12345":
        headers = ["#", "Question", "Model Answer"] + [f"Choice {i}" for i in range(1, 6)] + ["Correct Choice #"]
    else:
        headers = ["#", "Question", "Model Answer", "Choice A", "Choice B", "Choice C", "Choice D", "Choice E", "Correct Choice"]
    ws2.append(headers)
    for qid in sorted(choices_map.keys()):
        choices, correct = choices_map[qid]
        row = [qid, q_text.get(qid, ""), q_ans.get(qid, "")] + list(choices) + [correct]
        ws2.append(row)
        for cell in ws2[ws2.max_row]:
            cell.font = BODY_FONT
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = BORDER
    ws2.freeze_panes = "A2"
    wb.save(wb_path)
    print("OK", wb_path.name)


def main():
    for fname, (fmt, cmap) in CONFIG.items():
        p = BASE / fname
        if p.exists():
            rebuild_sheet(p, fmt, cmap)


if __name__ == "__main__":
    main()
