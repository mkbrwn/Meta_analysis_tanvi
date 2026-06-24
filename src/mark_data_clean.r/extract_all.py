#!/usr/bin/env python3
"""
extract_all.py
==============
Combined extraction of participant counts and proportion / rate data from
free-text columns in merged_all_sheets.csv.

Outputs (in src/mark_data_clean.r/):
    extracted_participants.csv   – study_number, total_participants, manual_clean
    extracted_outcomes.csv       – per-column numerator / denominator / proportion
    extraction_summary.csv       – success / failure summary per column
"""

import pandas as pd
import re
import os
import warnings

warnings.filterwarnings("ignore", category=FutureWarning)

# ── paths ──────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
BASE  = os.path.dirname(os.path.dirname(_HERE))

INPUT_CSV  = os.path.join(BASE, "data", "data_extraction_tanvi_050626",
                           "merged_all_sheets.csv")
OUTPUT_DIR = _HERE

STUDY_COL        = "study number"
PARTICIPANTS_COL = "number of participants"


# ════════════════════════════════════════════════════════════════════════════
#  1.  PARTICIPANT-COUNT EXTRACTION
# ════════════════════════════════════════════════════════════════════════════

def extract_total_participants(text: str):
    """
    Parse free-text from the 'number of participants' column.
    Returns (total: int or None, method: str, manual_clean: bool).
    """
    if not text or text.strip() == "":
        return None, "empty", False          # empty → nothing to clean

    t = text.strip()

    # ── Pre-check: age-like patterns (wrong column) ───────────────────────
    age_pats = [
        r'\d+\s*\(\d+[-–]\d+\)',             # "60 (49-75)"
        r'\d+\.?\d*\s*[±–-]\s*\d+\.?\d*',   # "63.4 ± 16.5"
        r'\d+\.?\d*\s*\(\s*\d+\.?\d*\s*\)', # "55.5 (15.9)"
        r'median', r'mean', r'IQR', r'range',
        r'≥?\d+\s*years', r'Age\b',
        r'\d+\.?\d+\s*\(SD', r'\d+\.?\d+\s*\(sd',
        r'\d+\s*\(Q[13]\s*;',
    ]
    for pat in age_pats:
        if re.search(pat, t, re.IGNORECASE):
            return None, "age_data_in_participants_col", True

    # ── helpers ───────────────────────────────────────────────────────────
    def _int(s):
        return int(s.replace(",", ""))

    # S1: plain integer
    m = re.match(r'^\s*(\d+)\s*$', t)
    if m:
        return int(m.group(1)), "simple_integer", False

    # S2: comma-separated number
    m = re.match(r'^\s*([\d,]+)\s*$', t)
    if m:
        return _int(m.group(1)), "comma_number", False

    # S3: leading "n = N"
    m = re.match(r'^\s*n\s*=\s*([\d,]+)', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "n_equals_start", False

    # S4: single "n = N" anywhere
    matches = re.findall(r'(?:^|[,;]\s*)n\s*=\s*([\d,]+)', t, re.IGNORECASE)
    if len(matches) == 1:
        return _int(matches[0]), "n_equals_single", False

    # S5: "X patients"
    m = re.search(r'(\d[\d,]*)\s+patients?', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "x_patients", False

    # S6: "X consecutive patients"
    m = re.search(r'(\d[\d,]*)\s+consecutive\s+patients?', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "x_consecutive_patients", False

    # S7: "N total"
    m = re.match(r'^\s*(\d[\d,]*)\s+total', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "x_total", False

    # S8: "total n=N"
    m = re.search(r'total\s+n\s*=\s*([\d,]+)', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "total_n_equals", False

    # S9: "n=N (" with subgroups
    m = re.search(r'n\s*=\s*([\d,]+)\s*\(', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "n_equals_with_subgroups", False

    # S10: "X cases"
    m = re.search(r'(\d[\d,]*)\s+cases?', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "x_cases", False

    # S11: "X study participants"
    m = re.search(r'(\d[\d,]*)\s+study\s+participants?', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "x_study_participants", False

    # S12: "A total of X"
    m = re.search(r'a\s+total\s+of\s+([\d,]+)', t, re.IGNORECASE)
    if m:
        return _int(m.group(1)), "a_total_of_x", False

    # S13: "30 (15+15)" → take leading number
    m = re.search(r'(\d[\d,]*)\s*\((\d+)\+?(\d+)\)', t)
    if m:
        return _int(m.group(1)), "x_with_groups", False

    # S14: "15+15"
    m = re.match(r'^\s*(\d+)\s*\+\s*(\d+)\s*$', t)
    if m:
        return int(m.group(1)) + int(m.group(2)), "sum_groups", False

    # S15: multiple "n=X" values → sum
    ns = re.findall(r'n\s*=\s*(\d+)', t, re.IGNORECASE)
    if len(ns) >= 2 and max(int(x) for x in ns) > 10:
        return sum(int(x) for x in ns), "sum_of_n_values", False

    # S16: leading number > 5
    m = re.match(r'^\s*(\d[\d,]*)\b', t)
    if m:
        v = _int(m.group(1))
        if v > 5:
            return v, "leading_number", False

    # S17: largest plausible number
    nums = [_int(x) for x in re.findall(r'\b(\d[\d,]*)\b', t)]
    cands = [n for n in nums if 10 <= n <= 200000]
    if cands:
        return max(cands), "largest_number_in_text", False

    return None, "unparseable", True


# ════════════════════════════════════════════════════════════════════════════
#  2.  PROPORTION / RATE EXTRACTION
# ════════════════════════════════════════════════════════════════════════════

def _back_calc(proportion, total_participants, method_prefix, is_mortality=False):
    """Back-calculate numerator/denominator from a proportion and participant count.

    manual_clean is TRUE only for mortality columns (inferred numerator).
    """
    if proportion is None or total_participants is None or total_participants <= 0:
        return {}
    num = min(round(proportion * total_participants), total_participants)
    return dict(numerator=num, denominator=total_participants,
                proportion=proportion,
                extraction_method=f"{method_prefix}_backcalc",
                manual_clean=is_mortality)


def extract_proportion(text: str, total_participants: int = None,
                       is_mortality: bool = False) -> dict:
    """Parse free-text describing a proportion / rate.

    Returns dict with: numerator, denominator, proportion,
                       extraction_method, manual_clean.

    - If only a proportion is found (no N/D), back-calculate using
      total_participants when available.
    - Empty cells → manual_clean = FALSE.
    """
    out = dict(numerator=None, denominator=None, proportion=None,
               extraction_method="unparseable", manual_clean=True)

    if not text or not str(text).strip():
        out.update(extraction_method="empty", manual_clean=False)
        return out

    t = str(text).strip()
    t = re.sub(r'^[\?！]\s*', '', t)
    t = re.sub(r'^NB:\s*', '', t, flags=re.IGNORECASE)
    t_clean = re.sub(r'%%', '%', t)                         # double-% typo
    t_clean = t_clean.replace('\u201c', '"').replace('\u201d', '"')
    t_clean = re.sub(r'"{1,}$', '', t_clean).strip()

    # S1: fraction N/D (possibly with n= prefix)
    frac = re.findall(r'(?:n\s*=\s*)?(\d[\d,]*)\s*/\s*(\d[\d,]*)',
                       t_clean, re.IGNORECASE)
    if frac:
        plausible = [(int(a.replace(',','')), int(b.replace(',','')))
                     for a, b in frac
                     if int(b.replace(',','')) > 0
                     and int(a.replace(',','')) <= int(b.replace(',',''))]
        if plausible:
            num, den = plausible[0]
            prop = num / den
            method = "fraction_compound" if len(plausible) > 1 else "fraction"
            out.update(numerator=num, denominator=den, proportion=prop,
                       extraction_method=method, manual_clean=False)
            # verify against parenthetical %
            pm = re.search(r'\((\d+\.?\d*)\s*%\)', t_clean)
            if pm and prop and abs(prop - float(pm.group(1))/100) > 0.05:
                out.update(manual_clean=True, extraction_method="fraction_pct_mismatch")
            return out

    # S2: decimal proportion  0.XXXX
    m = re.match(r'^\s*0\.(\d+)\s*$', t_clean)
    if m:
        val = float(t_clean)
        if 0 <= val <= 1.0:
            bc = _back_calc(val, total_participants, "decimal_proportion", is_mortality)
            out.update(bc) if bc else out.update(proportion=val,
                extraction_method="decimal_proportion", manual_clean=False)
            return out

    # S3: "N (X%)"
    m = re.match(r'^\s*(?:n\s*=\s*)?(\d[\d,]*)\s*[,/]?\s*\(\s*(\d+\.?\d*)\s*%\s*\)',
                 t_clean, re.IGNORECASE)
    if m:
        num = int(m.group(1).replace(',',''))
        pct = float(m.group(2))
        if 0 < pct <= 100:
            den = round(num / (pct / 100))
            out.update(numerator=num, denominator=den,
                       proportion=num/den if den else None,
                       extraction_method="count_with_pct", manual_clean=False)
            return out

    # S4: plain "N%"
    m = re.match(r'^\s*(\d+\.?\d*)\s*%\s*$', t_clean)
    if m:
        val = float(m.group(1)) / 100
        if 0 <= val <= 1.0:
            bc = _back_calc(val, total_participants, "pct_only", is_mortality)
            out.update(bc) if bc else out.update(proportion=val,
                extraction_method="pct_only", manual_clean=False)
            return out

    # S4b: "X% vs. Y%" – take first
    m = re.match(r'^\s*(\d+\.?\d*)\s*%\s*(?:vs\.?|/)\s*(\d+\.?\d*)\s*%?\s*$',
                 t_clean, re.IGNORECASE)
    if m:
        val = float(m.group(1)) / 100
        if 0 <= val <= 1.0:
            bc = _back_calc(val, total_participants, "pct_comparative", is_mortality)
            out.update(bc) if bc else out.update(proportion=val,
                extraction_method="pct_comparative_first", manual_clean=False)
            return out

    # S5: "n=N (X%)"
    m = re.match(r'^\s*n\s*=\s*(\d[\d,]*)\s*\(\s*(\d+\.?\d*)\s*%\s*\)',
                 t_clean, re.IGNORECASE)
    if m:
        num = int(m.group(1).replace(',',''))
        pct = float(m.group(2))
        if 0 < pct <= 100:
            den = round(num / (pct / 100))
            out.update(numerator=num, denominator=den,
                       proportion=num/den if den else None,
                       extraction_method="n_equals_with_pct", manual_clean=False)
            return out

    # S6: "n=N" only
    m = re.match(r'^\s*n\s*=\s*(\d[\d,]*)\s*$', t_clean, re.IGNORECASE)
    if m:
        num = int(m.group(1).replace(',',''))
        if total_participants and total_participants > 0:
            out.update(numerator=num, denominator=total_participants,
                       proportion=num/total_participants,
                       extraction_method="n_equals_only_backcalc",
                       manual_clean=is_mortality)
        else:
            out.update(numerator=num, extraction_method="n_equals_only",
                       manual_clean=True)
        return out

    # S7: explicit 100% / 0%
    if re.search(r'100\s*%', t_clean):
        bc = _back_calc(1.0, total_participants, "explicit_100pct", is_mortality)
        out.update(bc) if bc else out.update(proportion=1.0,
            extraction_method="explicit_100pct", manual_clean=False)
        return out
    if re.search(r'\b0\s*%\b', t_clean):
        bc = _back_calc(0.0, total_participants, "explicit_0pct", is_mortality)
        out.update(bc) if bc else out.update(proportion=0.0,
            extraction_method="explicit_0pct", manual_clean=False)
        return out

    # S8: descriptive text
    if re.search(r'implicit|implied|reports?\s|not (?:reported|documented|explicit)',
                 t_clean, re.IGNORECASE):
        out.update(extraction_method="descriptive_text", manual_clean=True)
        return out

    # S9: standalone 0 or 1
    m = re.match(r'^\s*(\d+\.?\d*)\s*$', t_clean)
    if m:
        val = float(m.group(1))
        if val in (0, 1):
            bc = _back_calc(val, total_participants,
                            "standalone_zero" if val == 0 else "standalone_one",
                            is_mortality)
            out.update(bc) if bc else out.update(proportion=val,
                extraction_method="standalone_zero" if val == 0 else "standalone_one",
                manual_clean=False)
            return out
        out.update(numerator=val, extraction_method="standalone_number",
                   manual_clean=True)
        return out

    return out   # unparseable


# ════════════════════════════════════════════════════════════════════════════
#  3.  HELPER: build raw-text map per column
# ════════════════════════════════════════════════════════════════════════════

def build_raw_text_map(df, col):
    out = {}
    for sn, grp in df.groupby(STUDY_COL):
        vals = grp[col].dropna().astype(str).tolist()
        vals = [v.strip() for v in vals
                if v.strip() and v.strip().lower() not in ("nan", "none", "")]
        out[sn] = " | ".join(vals) if vals else ""
    return out


# ════════════════════════════════════════════════════════════════════════════
#  4.  OUTCOME COLUMNS
# ════════════════════════════════════════════════════════════════════════════

TARGET_COLUMNS = [
    'icu mortality',
    '28d mortality',
    '30d mortality',
    '60d mortality',
    '90d mortality',
    'hospital mortality',
    'mechnical ventilation (proportion requiring)',
    'niv (proportion requiring)',
    'ecmo (proportion requiring)',
    'rrt (proportion requiring)',
    'vasopressor / inotropic support (proportion requiring)',
    'septic shock (proportion with)',
    'respiratory failure (proportion with)',
    'ards (proportion with)',
    'copd',
    'smokers',
    'diabetes',
]


# ════════════════════════════════════════════════════════════════════════════
#  5.  MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    print("Loading data …")
    df = pd.read_csv(INPUT_CSV, low_memory=False)
    study_numbers = sorted(df[STUDY_COL].dropna().unique())
    print(f"  {len(df)} rows, {len(study_numbers)} unique study numbers\n")

    # ── 5a. Extract participant counts ────────────────────────────────────
    print("── Step 1: Extracting participant counts ──")
    part_rows = []
    for sn, grp in df.groupby(STUDY_COL):
        texts = grp[PARTICIPANTS_COL].dropna().astype(str).tolist()
        texts = [t.strip() for t in texts
                 if t.strip() and t.strip().lower() not in ("nan", "none", "")]
        raw = " | ".join(texts) if texts else ""
        total, method, mc = extract_total_participants(raw)
        part_rows.append(dict(study_number=sn, raw_participants_text=raw,
                              total_participants=total,
                              extraction_method=method, manual_clean=mc))

    part_df = pd.DataFrame(part_rows).sort_values("study_number").reset_index(drop=True)
    part_ok  = (~part_df["manual_clean"]).sum()
    part_man = part_df["manual_clean"].sum()
    print(f"  Extracted: {part_ok}/{len(part_df)}  |  manual_clean: {part_man}/{len(part_df)}")

    part_csv = os.path.join(OUTPUT_DIR, "extracted_participants.csv")
    part_df.to_csv(part_csv, index=False)
    print(f"  Saved: {part_csv}\n")

    # Build lookup map
    part_map = dict(zip(part_df["study_number"],
                        part_df["total_participants"].fillna(0).astype(int)))

    # ── 5b. Extract outcome columns ───────────────────────────────────────
    print("── Step 2: Extracting outcome columns ──")
    all_results = {}

    for col in TARGET_COLUMNS:
        if col not in df.columns:
            print(f"  ⚠  Column not found: '{col}'")
            continue

        raw_map = build_raw_text_map(df, col)
        rows, n_ok, n_man = [], 0, 0
        mortality_col = "mortality" in col.lower()

        for sn in study_numbers:
            raw  = raw_map.get(sn, "")
            tp   = part_map.get(sn, None)
            if tp is not None and tp <= 0:
                tp = None
            p = extract_proportion(raw, total_participants=tp,
                                   is_mortality=mortality_col)
            rows.append(dict(
                **{STUDY_COL: sn},
                **{f"raw_{col}":                raw},
                **{f"{col}_numerator":          p["numerator"]},
                **{f"{col}_denominator":        p["denominator"]},
                **{f"{col}_proportion":         p["proportion"]},
                **{f"{col}_method":             p["extraction_method"]},
                **{f"{col}_manual_clean":       p["manual_clean"]},
            ))
            n_man += p["manual_clean"]
            n_ok  += not p["manual_clean"]

        col_df = pd.DataFrame(rows)
        all_results[col] = col_df
        print(f"  ✓ {col:55s}  extracted: {n_ok}/{len(study_numbers)}"
              f"  |  manual_clean: {n_man}/{len(study_numbers)}")

    # ── 5c. Assemble wide output (participants + outcomes) ────────────────
    print("\nAssembling wide output …")
    wide = part_df.rename(columns={"study_number": STUDY_COL})[[STUDY_COL, "total_participants", "manual_clean"]].copy()
    wide = wide.rename(columns={
        "total_participants": "participants_total",
        "manual_clean":       "participants_manual_clean",
    })
    for col, cdf in all_results.items():
        wide = wide.merge(cdf, on=STUDY_COL, how="left")
    wide.sort_values(STUDY_COL, inplace=True)
    wide.reset_index(drop=True, inplace=True)

    out_csv = os.path.join(OUTPUT_DIR, "extracted_counts.csv")
    wide.to_csv(out_csv, index=False)
    print(f"  Saved: {out_csv}  ({wide.shape[0]} studies × {wide.shape[1]} columns)")

    # ── 5d. Summary ───────────────────────────────────────────────────────
    print("\n── Manual-clean summary ──")
    summary_rows = []
    for col in TARGET_COLUMNS:
        mc = f"{col}_manual_clean"
        mt = f"{col}_method"
        if mc in wide.columns:
            n_total = int(wide[mc].count())
            n_man   = int(wide[mc].sum())
            summary_rows.append(dict(column=col, total_studies=n_total,
                                     extracted_ok=n_total - n_man,
                                     manual_clean_needed=n_man))
            print(f"  {col:55s}  {n_total - n_man:3d} ok  |  {n_man:3d} manual")
    pd.DataFrame(summary_rows).to_csv(
        os.path.join(OUTPUT_DIR, "extraction_summary.csv"), index=False)


if __name__ == "__main__":
    main()
