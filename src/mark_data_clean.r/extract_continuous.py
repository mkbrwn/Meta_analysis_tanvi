#!/usr/bin/env python3
"""
extract_continuous.py
=====================
Extract continuous summary statistics (median, IQR, mean, SD) from free-text
columns in merged_all_sheets.csv.

Target columns: age, icu length of stay, hospital length of stay.

For each study and each variable the script produces:
    median, iqr_lower, iqr_upper, mean, sd,
    extraction_method, manual_clean

Output:  src/mark_data_clean.r/extracted_continuous.csv
"""

import pandas as pd
import re
import os

# ── paths ──────────────────────────────────────────────────────────────────
_HERE = os.path.dirname(os.path.abspath(__file__))
BASE  = os.path.dirname(os.path.dirname(_HERE))
INPUT_CSV  = os.path.join(BASE, "data", "data_extraction_tanvi_050626",
                           "merged_all_sheets.csv")
OUTPUT_DIR = _HERE
STUDY_COL  = "study number"

# Columns to extract
TARGET_COLUMNS = {
    "age":                                    "age",
    "icu length of stay":                     "icu_los",
    "hospital length of stay":                "hosp_los",
}


# ════════════════════════════════════════════════════════════════════════════
#  NUMBER HELPERS
# ════════════════════════════════════════════════════════════════════════════

def _num(s):
    """Parse a string to float, stripping commas."""
    return float(s.replace(",", "").replace(" ", ""))


def _safe_float(s):
    try:
        return _num(s)
    except (ValueError, TypeError):
        return None


# ════════════════════════════════════════════════════════════════════════════
#  EXTRACTION FUNCTION
# ════════════════════════════════════════════════════════════════════════════

def extract_continuous_stats(text: str) -> dict:
    """Parse free-text describing a continuous variable and return a dict with:
        median, iqr_lower, iqr_upper, mean, sd, extraction_method, manual_clean

    Returns the FIRST overall-cohort statistics found (ignores subgroup
    comparisons when an overall value is present).
    """
    out = dict(median=None, iqr_lower=None, iqr_upper=None,
               mean=None, sd=None,
               extraction_method="unparseable", manual_clean=True)

    if not text or not str(text).strip():
        out.update(extraction_method="empty", manual_clean=False)
        return out

    t = str(text).strip()

    # ── pre-processing ────────────────────────────────────────────────────
    # normalise unicode
    t = t.replace('\u00b1', '±').replace('\u2212', '-').replace('\u2013', '-')
    t = t.replace('\u2014', '-')  # em-dash → hyphen
    t = t.replace('–', '-')       # en-dash → hyphen
    t = re.sub(r'\+/-', '±', t)
    t = re.sub(r'\+/\-', '±', t)

    # ── if there are multiple groups separated by ";", prefer the first ───
    # But only if the first segment looks like it has stats.
    # e.g. "survivor group (n=49): 70.0 ± 2.0; non-survivor group ..."
    segments = re.split(r';\s*(?:non-?survivor|control|placebo|group\s*[234])',
                        t, flags=re.IGNORECASE)
    t_work = segments[0].strip()

    # ── Strategy 1: "median ... (IQR ...)" or "median ... [IQR]" ──────────
    # Patterns seen:
    #   "median 13 days (IQR 8–23)"
    #   "median [IQR] 68 [56, 80]"
    #   "median (IQR) in days: 5 (3, 9)"
    #   "days median (IQR) 13 (7, 24)"
    #   "Median (IQR) in days: no ARDS = 5 (3, 9); ARDS = 9 (5, 15)"
    #   "12 (8 ; 18)"
    #   "6 (4-9)"
    #   "70.0 ± 2.0" (mean ± sd)

    # Try to find median + IQR together first
    # Pattern A: median VALUE (IQR Q1–Q3) or median VALUE [Q1, Q3]
    m = re.search(
        r'median\s+(?:\w+\s+)?(?:\(IQR\)\s*(?:in\s+\w+:?\s*)?)?'
        r'(\d+\.?\d*)\s*[(\[]\s*(\d+\.?\d*)\s*[,;–\-]\s*(\d+\.?\d*)\s*[)\]]',
        t_work, re.IGNORECASE)
    if m:
        med, q1, q3 = _safe_float(m.group(1)), _safe_float(m.group(2)), _safe_float(m.group(3))
        # Also try to find mean ± SD in the same text
        mean, sd = _find_mean_sd(t_work)
        method = "median_iqr_with_mean_sd" if mean is not None else "median_iqr"
        out.update(median=med, iqr_lower=q1, iqr_upper=q3,
                   mean=mean, sd=sd, extraction_method=method, manual_clean=False)
        return out

    # Pattern B: "N (Q1–Q3)" or "N [Q1, Q3]" without explicit "median" keyword
    # e.g. "13 (7, 24)", "6 (4-9)", "68 [56, 80]", "12 (8 ; 18)"
    m = re.search(
        r'(\d+\.?\d*)\s*[(\[]\s*(\d+\.?\d*)\s*[,;–\-]\s*(\d+\.?\d*)\s*[)\]]',
        t_work)
    if m:
        val, q1, q3 = _safe_float(m.group(1)), _safe_float(m.group(2)), _safe_float(m.group(3))
        if q1 is not None and q3 is not None and q1 < val < q3:
            mean, sd = _find_mean_sd(t_work)
            method = "iqr_parenthetical_with_mean_sd" if mean is not None else "iqr_parenthetical"
            out.update(median=val, iqr_lower=q1, iqr_upper=q3,
                       mean=mean, sd=sd, extraction_method=method, manual_clean=False)
            return out

    # ── Strategy 2: mean ± SD ─────────────────────────────────────────────
    # Patterns: "42.3 ± 14.7", "75.03± 13.41", "18.7+/-9.6", "9.2 ± 11 days"
    mean, sd = _find_mean_sd(t_work)
    if mean is not None:
        # Also try to find median/IQR in the same text
        med, q1, q3 = _find_median_iqr(t_work)
        method = "mean_sd_with_median_iqr" if med is not None else "mean_sd"
        out.update(median=med, iqr_lower=q1, iqr_upper=q3,
                   mean=mean, sd=sd, extraction_method=method, manual_clean=False)
        return out

    # ── Strategy 3: median keyword with plain number ──────────────────────
    # e.g. "median 13 days", "median age of 73 years"
    m = re.search(r'median\s+(?:age\s+(?:of\s+)?)?(?:\w+\s+)?(\d+\.?\d*)',
                  t_work, re.IGNORECASE)
    if m:
        med = _safe_float(m.group(1))
        if med is not None:
            out.update(median=med, extraction_method="median_only", manual_clean=False)
            return out

    # ── Strategy 4: mean keyword with plain number ────────────────────────
    # e.g. "Mean age was 60.5 years", "mean 59.5 ± 16.6" (already caught above)
    m = re.search(r'mean\s+(?:age\s+(?:was\s+|of\s+)?)?(?:\w+\s+)?(\d+\.?\d*)',
                  t_work, re.IGNORECASE)
    if m:
        mean_val = _safe_float(m.group(1))
        if mean_val is not None:
            # Check for SD nearby: "mean X (SD Y)" or "mean X; SD Y"
            sd_m = re.search(r'\(?\s*SD\s+(\d+\.?\d*)\s*\)?', t_work, re.IGNORECASE)
            sd_val = _safe_float(sd_m.group(1)) if sd_m else None
            out.update(mean=mean_val, sd=sd_val,
                       extraction_method="mean_only" if sd_val is None else "mean_sd",
                       manual_clean=False)
            return out

    # ── Strategy 5: "median ... Q1; Q3" pattern ───────────────────────────
    # e.g. "83 (81; 85)"
    m = re.search(r'(\d+\.?\d*)\s*\(\s*(\d+\.?\d*)\s*;\s*(\d+\.?\d*)\s*\)', t_work)
    if m:
        val, q1, q3 = _safe_float(m.group(1)), _safe_float(m.group(2)), _safe_float(m.group(3))
        if q1 is not None and q3 is not None and q1 < val < q3:
            out.update(median=val, iqr_lower=q1, iqr_upper=q3,
                       extraction_method="median_iqr_semicolon", manual_clean=False)
            return out

    # ── Strategy 6: "mean (SD) X.X (Y.Y)" ─────────────────────────────────
    # e.g. "Age (years), mean (SD) 55.7 (16.8)"
    #       "mean (SD) age was 67.5 (12.4) years"
    #       "Mean (SD): 55.5 (15.9)"
    #       "mean (SD): 75.03 ± 13.41"  (already caught by ± strategy)
    m = re.search(
        r'mean\s*\(SD\)\s*(?:[:\w\s]*?)\s*(\d+\.?\d*)\s*\(\s*(\d+\.?\d*)\s*\)',
        t_work, re.IGNORECASE)
    if m:
        mean_val, sd_val = _safe_float(m.group(1)), _safe_float(m.group(2))
        if mean_val is not None:
            out.update(mean=mean_val, sd=sd_val,
                       extraction_method="mean_sd_parenthetical", manual_clean=False)
            return out

    # ── Strategy 7: "X.X (Y.Y)" mean(SD) when X.X > Y.Y ─────────────────
    # e.g. "62.88 (18.75)", "60.9 (16.07)", "61.6 (17.6)"
    m = re.match(r'^\s*(\d+\.?\d*)\s*\(\s*(\d+\.?\d*)\s*\)\s*$', t_work)
    if m:
        val1, val2 = _safe_float(m.group(1)), _safe_float(m.group(2))
        if val1 is not None and val2 is not None and val1 > val2:
            out.update(mean=val1, sd=val2,
                       extraction_method="mean_sd_implicit", manual_clean=False)
            return out

    # ── Strategy 8: "mean +/- SD N (+/- M)" ──────────────────────────────
    # e.g. "mean +/- SD 60 (+/- 16.5)"
    m = re.search(r'mean\s*±\s*SD\s+(\d+\.?\d*)\s*\(?\s*±\s*(\d+\.?\d*)\)?',
                  t_work, re.IGNORECASE)
    if m:
        mean_val, sd_val = _safe_float(m.group(1)), _safe_float(m.group(2))
        if mean_val is not None:
            out.update(mean=mean_val, sd=sd_val,
                       extraction_method="mean_sd_explicit", manual_clean=False)
            return out

    # ── Strategy 9: "mean IQR: X (Y)" ────────────────────────────────────
    # e.g. "Mean IQR: empiricial antiviral 56.1 (14.5)"
    m = re.search(r'mean\s+IQR\s*:?\s*\S+\s+(\d+\.?\d*)\s*\(\s*(\d+\.?\d*)\s*\)',
                  t_work, re.IGNORECASE)
    if m:
        mean_val, sd_val = _safe_float(m.group(1)), _safe_float(m.group(2))
        if mean_val is not None:
            out.update(mean=mean_val, sd=sd_val,
                       extraction_method="mean_iqr_label", manual_clean=False)
            return out

    # ── Strategy 10: plain number (may be mean or median, ambiguous) ──────
    m = re.match(r'^\s*(\d+\.?\d*)\s*(?:days?|years?)?\s*$', t_work, re.IGNORECASE)
    if m:
        val = _safe_float(m.group(1))
        if val is not None:
            out.update(mean=val, extraction_method="plain_number", manual_clean=True)
            return out

    # ── Strategy 11: "X-Y" range → midpoint as mean ──────────────────────
    m = re.match(r'^\s*(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)\s*$', t_work)
    if m:
        lo, hi = _safe_float(m.group(1)), _safe_float(m.group(2))
        if lo is not None and hi is not None and lo < hi:
            out.update(mean=(lo + hi) / 2, extraction_method="range_midpoint",
                       manual_clean=True)
            return out

    # Nothing matched
    return out


def _find_mean_sd(text):
    """Extract mean ± SD from text. Returns (mean, sd) or (None, None)."""
    # "N.N ± N.N" or "N.N +/- N.N" or "N.N±N.N"
    m = re.search(r'(\d+\.?\d*)\s*±\s*(\d+\.?\d*)', text)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2))
    m = re.search(r'(\d+\.?\d*)\s*\+/-\s*(\d+\.?\d*)', text)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2))
    # "(SD N.N)" after a number: "60.5 years (SD 16.4)"
    m = re.search(r'(\d+\.?\d*)\s*(?:years?|days?)?\s*\(SD\s+(\d+\.?\d*)\)',
                  text, re.IGNORECASE)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2))
    # "mean X; SD Y"
    m = re.search(r'mean\s+(\d+\.?\d*)\s*;?\s*SD\s+(\d+\.?\d*)',
                  text, re.IGNORECASE)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2))
    # "mean years X.X; SD Y.Y"
    m = re.search(r'mean\s+years?\s+(\d+\.?\d*)\s*;?\s*SD\s+(\d+\.?\d*)',
                  text, re.IGNORECASE)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2))
    return None, None


def _find_median_iqr(text):
    """Extract median and IQR from text. Returns (median, q1, q3) or (None, None, None)."""
    # "median X (IQR Y–Z)" or "median X (Y-Z)"
    m = re.search(r'median\s+\w*\s*(\d+\.?\d*)\s*\(IQR\s+(\d+\.?\d*)\s*[-–]\s*(\d+\.?\d*)\)',
                  text, re.IGNORECASE)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2)), _safe_float(m.group(3))
    # "median [IQR] X [Y, Z]"
    m = re.search(r'median\s+\[IQR\]\s*(\d+\.?\d*)\s*\[\s*(\d+\.?\d*)\s*,\s*(\d+\.?\d*)\s*\]',
                  text, re.IGNORECASE)
    if m:
        return _safe_float(m.group(1)), _safe_float(m.group(2)), _safe_float(m.group(3))
    return None, None, None


# ════════════════════════════════════════════════════════════════════════════
#  MAIN
# ════════════════════════════════════════════════════════════════════════════

def main():
    print("Loading data …")
    df = pd.read_csv(INPUT_CSV, low_memory=False)
    study_numbers = sorted(df[STUDY_COL].dropna().unique())
    print(f"  {len(df)} rows, {len(study_numbers)} unique study numbers\n")

    # ── process each target column ─────────────────────────────────────────
    all_results = {}

    for col, prefix in TARGET_COLUMNS.items():
        if col not in df.columns:
            print(f"  ⚠  Column not found: '{col}'")
            continue

        # Build raw text map
        raw_map = {}
        for sn, grp in df.groupby(STUDY_COL):
            vals = grp[col].dropna().astype(str).tolist()
            vals = [v.strip() for v in vals
                    if v.strip() and v.strip().lower() not in ("nan", "none", "")]
            raw_map[sn] = " | ".join(vals) if vals else ""

        rows = []
        n_ok = 0
        n_man = 0

        for sn in study_numbers:
            raw = raw_map.get(sn, "")
            parsed = extract_continuous_stats(raw)

            rows.append({
                STUDY_COL:          sn,
                f"raw_{prefix}":    raw,
                f"{prefix}_median":      parsed["median"],
                f"{prefix}_iqr_lower":   parsed["iqr_lower"],
                f"{prefix}_iqr_upper":   parsed["iqr_upper"],
                f"{prefix}_mean":        parsed["mean"],
                f"{prefix}_sd":          parsed["sd"],
                f"{prefix}_method":      parsed["extraction_method"],
                f"{prefix}_manual_clean": parsed["manual_clean"],
            })
            n_man += parsed["manual_clean"]
            n_ok  += not parsed["manual_clean"]

        col_df = pd.DataFrame(rows)
        all_results[prefix] = col_df
        print(f"  ✓ {col:30s} → {prefix:10s}  extracted: {n_ok}/{len(study_numbers)}"
              f"  |  manual_clean: {n_man}/{len(study_numbers)}")

    # ── assemble wide output ───────────────────────────────────────────────
    print("\nAssembling output …")
    wide = pd.DataFrame({STUDY_COL: study_numbers})
    for prefix, cdf in all_results.items():
        wide = wide.merge(cdf, on=STUDY_COL, how="left")
    wide.sort_values(STUDY_COL, inplace=True)
    wide.reset_index(drop=True, inplace=True)

    out_csv = os.path.join(OUTPUT_DIR, "extracted_continuous.csv")
    wide.to_csv(out_csv, index=False)
    print(f"  Saved: {out_csv}  ({wide.shape[0]} studies × {wide.shape[1]} columns)")

    # ── summary ────────────────────────────────────────────────────────────
    print("\n── Summary ──")
    for prefix in TARGET_COLUMNS.values():
        mc = f"{prefix}_manual_clean"
        if mc in wide.columns:
            n_total = int(wide[mc].count())
            n_man   = int(wide[mc].sum())
            n_has_med = wide[f"{prefix}_median"].notna().sum()
            n_has_mn  = wide[f"{prefix}_mean"].notna().sum()
            print(f"  {prefix:10s}  median: {int(n_has_med):3d}  mean: {int(n_has_mn):3d}"
                  f"  manual: {n_man}/{n_total}")


if __name__ == "__main__":
    main()
