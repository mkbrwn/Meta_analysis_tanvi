# Jesus Is LORD

# from claude (v3)

# slight adaptation

# =============================================================================
# Age Data Parsing Script — Version 3
# =============================================================================
# Fixes applied versus v2 (by row number in the data):
#
#  2  - "years mean +/- SD 75.1 ± 14.0": now correctly reads 75.1 as mean,
#         14.0 as SD. Previously SD regex matched 75.1.
#  6  - "mean years 65.4; SD 14.6; median years 67": semicolons here separate
#         values not groups. Now detected as overall, not subgrouped.
#  10 - "≥18 years: ... KP-CAP 68 [...]; SP-CAP 73 [...]": category prefix
#         stripped before group splitting so KP-CAP/SP-CAP become group names.
#  14 - "(Q1 ; Q3) 83 (81; 85)": Q1;Q3 semicolon no longer triggers subgrouping.
#  15 - ">18: ... no LP 65; LP 59": category prefix stripped before splitting.
#  20 - "≥18 years: positive BC ...; negative BC ...": category prefix stripped.
#  22 - "62.88 (18.75)": single bracketed number now treated as SD.
#  50 - "Median 63 years (SOC) vs 59 years (CPA)": group_1_name now "SOC"/"CPA"
#         not "Median".
#  56 - "Psittacosis 48 years vs Legionella 60 years": group_1_name now
#         "Psittacosis" not "Median age".
#  72 - "... overall; 48 (36-55) indigenous; 64 (57-70) non-indigenous":
#         subgroup median/IQR type inherited from overall when not restated.
#  75 - "median iqr - septic shock 63 (26), noss 60 (30)": now detected as
#         2 groups split on comma after a dash pattern.
#  79 - "mean (SD) 55.7 (16.8)": SD now pulled (16.8).
#  81 - "range 18 to 101, IQR 51 to 71": both range and IQR now captured.
#  82 - "nonsurvivors ... 56.5 + 7.8 ... survivors ... 48 + 11.3": group names
#         now "nonsurvivors" / "survivors"; averages pulled correctly.
#  84 - "mean (SD) age was 67.5 (12.4)": SD now pulled (12.4).
#  86 - "median (range) age 63.8 (23-80)": range now pulled.
#  89 - "ranged from 14 to 71 ... mean of 35.27": range now pulled alongside mean.
#  90 - "bQ (n=104) 54 (44-67) bM (n=106) 55 (43-67)": detected as 2 groups.
#  92 - "Age, years (median [IQR]) 62 [46; 76]": no longer split as subgroup;
#         IQR now pulled.
#  93 - "Mean (SD): 55.5 (15.9)": SD now pulled (15.9).
#  95 - "58-62": bare hyphenated range now pulled as unspecified.
#  102 - "median (IQR) 63.4 ± 16.5": IQR keyword now takes priority; ± value
#          recorded as dispersion field rather than overriding to SD.
#  104 - "mean age of 45.5 ± 14.5 (median, 47)": type_of_average now correctly
#          "mean"; "(median, 47)" is a secondary stat, not the primary type.
#  106 - "41.6 ± 11.9 (range 21 to 61)": range now also captured.
#  108 - "Mean IQR: empiricial antiviral 56.1 (14.5). No antiviral 57.6 (16.3)":
#          group_1_name now "empiricial antiviral"; IQR keyword inherited.
#  6,14,92 - no longer incorrectly subgrouped.
#  108,56,6,10,15,20 - group_1_name now correct.
#  82 - group_2_name now correct.
#
# =============================================================================

library(dplyr)
library(stringr)

#my ammendment
library(readxl)
age <- read_xlsx("C:/Users/goddab/Downloads/age.xlsx")

# =============================================================================
# HELPER: strip_category_prefix()
# =============================================================================
# PURPOSE: Many cells begin with a category prefix like "≥18 years: " or
#          ">18: years median IQR; " before the actual data. When these cells
#          contain multiple groups separated by semicolons, the prefix was
#          incorrectly being treated as the first group.
#          This function removes that leading category prefix before splitting.
# INPUT:   Full cell text
# OUTPUT:  Text with leading category prefix removed (if present)

strip_category_prefix <- function(text) {
  if (is.na(text)) return(text)
  
  # Pattern: optional ≥/>/= sign, digits, "years"/"y", optional further words,
  # then a colon — this is the category prefix pattern
  # e.g. "≥18 years: ", ">18: years median IQR; ", "≥80 years: year median "
  # We strip everything up to and including the FIRST colon that follows this pattern
  stripped <- str_remove(text, regex(
    "^[>=<≥≤]?\\s*\\d+\\s*y(ears?)?\\s*:?\\s*(years?|year|y)?\\s*,?\\s*",
    ignore_case = TRUE
  ))
  
  # Also strip a leading format descriptor that ends with a semicolon before data
  # e.g. "years median IQR; " at the start (after the category prefix above)
  # But only strip if what remains still has numeric content
  stripped2 <- str_remove(stripped, regex(
    "^(years?|year|y)?\\s*(mean|median|average)?\\s*(\\+/?-|±)?\\s*(SD|IQR|\\(IQR\\))?\\s*;?\\s*",
    ignore_case = TRUE
  ))
  
  # Only use stripped2 if it still contains a digit (don't over-strip)
  if (str_detect(stripped2, "\\d")) {
    return(str_trim(stripped2))
  } else if (str_detect(stripped, "\\d")) {
    return(str_trim(stripped))
  } else {
    return(str_trim(text))
  }
}


# =============================================================================
# HELPER: is_value_semicolon()
# =============================================================================
# PURPOSE: Determines whether semicolons in a cell are separating VALUES
#          (e.g. "mean 65.4; SD 14.6; median 67") rather than GROUPS.
#          This prevents those cells from being incorrectly split into subgroups.
# INPUT:   Cell text
# OUTPUT:  TRUE if semicolons are value separators (not group separators)

is_value_semicolon <- function(text) {
  if (!str_detect(text, ";")) return(FALSE)
  
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  
  # A cell uses value-semicolons if segments are short fragments that each
  # contain ONLY a stat keyword or a single number — not a full age statement
  # Signs of value-semicolons:
  #   - A segment is just "SD 14.6" or "median years 67"
  #   - A segment contains only a keyword (mean/median/SD/IQR) + optional number
  #   - Segments contain "Q1" or "Q3" (IQR notation)
  value_like <- sapply(segs, function(s) {
    str_detect(s, regex(
      "^(SD|IQR|Q[123]|mean|median|average|range|\\d+\\.?\\d*|years?|y)\\b",
      ignore_case = TRUE
    )) && !str_detect(s, regex("[a-z]{4,}\\s+\\d", ignore_case = TRUE))
    # The second condition: if a segment has a long word followed by a number
    # it is more likely a group name + value, not a value fragment
  })
  
  # Also flag Q1/Q3 pattern specifically
  has_q1q3 <- str_detect(text, regex("Q1\\s*;\\s*Q3", ignore_case = TRUE))
  
  return(all(value_like) || has_q1q3)
}


# =============================================================================
# PART 1: parse_age_block()
# =============================================================================
# PURPOSE: Extracts all age-related values from ONE block of text.
# INPUT:   A character string for a single group/cohort
# OUTPUT:  A named list of 12 values (added range alongside IQR for row 81/106)

parse_age_block <- function(text, inherit_type_avg = NULL, inherit_disp_type = NULL) {
  
  # Arguments inherit_type_avg and inherit_disp_type allow the overall measure
  # type (e.g. "median", "IQR") to be passed into subgroup blocks when the
  # subgroup text does not restate those keywords (e.g. row 72).
  
  out <- list(
    category           = NA_character_,
    average_age        = NA_real_,
    type_of_average    = NA_character_,
    measure_dispersion = NA_character_,
    dispersion         = NA_real_,
    LQR                = NA_real_,
    UQR                = NA_real_,
    range_lower        = NA_real_,
    range_upper        = NA_real_,
    lower_unspecified  = NA_real_,
    upper_unspecified  = NA_real_
  )
  
  if (is.na(text) || str_trim(text) == "") return(out)
  
  text <- str_trim(text)
  text <- str_replace_all(text, "\u2013|\u2014", "-")  # normalise dashes
  text <- str_replace_all(text, "\u00b1", "±")          # normalise ±
  text <- str_replace_all(text, "\u2265", ">=")         # normalise ≥
  text <- str_replace_all(text, "\u2264", "<=")         # normalise ≤
  
  # ---------------------------------------------------------------------------
  # CATEGORY
  # ---------------------------------------------------------------------------
  cat_match <- str_extract(text, regex(
    paste0(
      "adults?\\s*(>=?\\s*\\d+\\s*y(ears?)?)?|",
      "[>=<]+\\s*\\d+\\s*y(ears?)?|",
      "\\d+\\s*-\\s*\\d+\\s*y(ears?)?|",
      "grouped ages?[^;]*"
    ),
    ignore_case = TRUE
  ))
  out$category <- if (!is.na(cat_match)) str_trim(cat_match) else NA_character_
  
  # ---------------------------------------------------------------------------
  # TYPE OF AVERAGE
  # ---------------------------------------------------------------------------
  # FIX (row 104): "mean age of 45.5 (median, 47)" — "median" here is a
  # secondary statistic. We check for mean FIRST; only fall through to median
  # if mean is absent. This correctly gives mean priority.
  #
  # FIX (row 104 cont.): parenthesised "(median, number)" pattern excluded
  # from triggering median detection when mean is also present.
  
  has_mean   <- str_detect(text, regex("\\bmean\\b|\\baverage\\b", ignore_case = TRUE))
  # Exclude "(median, X)" as a secondary stat marker
  text_for_median <- str_remove(text, regex("\\(\\s*median\\s*,\\s*\\d+\\.?\\d*\\s*\\)", ignore_case = TRUE))
  has_median <- str_detect(text_for_median, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))
  
  if (has_mean) {
    out$type_of_average <- "mean"
  } else if (has_median) {
    out$type_of_average <- "median"
  } else if (!is.null(inherit_type_avg)) {
    out$type_of_average <- inherit_type_avg  # inherit from overall if not stated
  } else {
    out$type_of_average <- "unknown"
  }
  
  # ---------------------------------------------------------------------------
  # AVERAGE AGE
  # ---------------------------------------------------------------------------
  avg <- NA_real_
  
  # FIX (row 2): "years mean +/- SD 75.1 ± 14.0"
  # Previously the SD regex matched 75.1 because "+/- SD 75.1" triggered it.
  # Now: we specifically look for mean/median followed by a number,
  # excluding numbers that appear right after SD keywords.
  
  # Pattern A: mean/median/average keyword followed by number
  m <- str_extract(text, regex(
    "(?:mean\\s*(?:age(?:d|s)?)?(?:\\s*(?:was|of|at admission of))?|median\\s*(?:age(?:\\s*(?:of|,|was))?)?|average\\s*(?:age\\s*was)?)\\s*[~=:]?\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  if (!is.na(m)) avg <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))
  
  # Pattern B: "Mean 75", "Mean 49.5y" — keyword then immediate number
  if (is.na(avg)) {
    m2 <- str_extract(text, regex("(?:mean|median|average)\\s+(\\d+\\.?\\d*)\\s*y?\\b", ignore_case = TRUE))
    if (!is.na(m2)) avg <- as.numeric(str_extract(m2, "\\d+\\.?\\d*"))
  }
  
  # Pattern C: "63 (mean)" — number before the word mean in brackets
  if (is.na(avg)) {
    m3 <- str_extract(text, regex("(\\d+\\.?\\d*)\\s*\\(\\s*mean\\s*\\)", ignore_case = TRUE))
    if (!is.na(m3)) avg <- as.numeric(str_extract(m3, "^\\d+\\.?\\d*"))
  }
  
  # Pattern D: "~65" — approximate number
  if (is.na(avg)) {
    m4 <- str_extract(text, "~\\s*(\\d+\\.?\\d*)")
    if (!is.na(m4)) avg <- as.numeric(str_extract(m4, "\\d+\\.?\\d*"))
  }
  
  # Pattern E: bare number (possibly followed by y) before any bracket or ± or end
  # Fallback for cells like "52.91", "62.88", "72.9"
  if (is.na(avg)) {
    m5 <- str_extract(text, "^[^\\d]*(\\d+\\.?\\d*)\\s*y?\\s*(?:[\\(\\[±\\+]|$)")
    if (!is.na(m5)) avg <- as.numeric(str_extract(m5, "\\d+\\.?\\d*"))
  }
  
  # Pattern F: "58-62" bare hyphenated range with no bracket — treat lower as average
  # (captured separately as range; average left NA in this case)
  
  out$average_age <- avg
  
  # ---------------------------------------------------------------------------
  # DISPERSION
  # FIX (rows 79, 84, 93): "mean (SD) 55.7 (16.8)" / "Mean (SD): 55.5 (15.9)"
  #   The bracketed number after an explicit (SD) label was not being captured.
  #   New pattern: "(SD) X" or "(SD): X" or number in brackets following (SD) label.
  #
  # FIX (row 102): "median (IQR) 63.4 ± 16.5"
  #   IQR keyword present but ± was triggering SD. Now IQR takes priority if
  #   the keyword is explicitly present, and the ± value is stored as dispersion.
  #
  # FIX (rows 81, 106): cells with BOTH range and IQR — both are now captured.
  #
  # FIX (row 22): single number in brackets with no keyword -> treat as SD.
  #
  # FIX (row 86): "median (range) age 63.8 (23-80)" — range keyword in brackets
  #   before "age" was not matched. Pattern now handles this.
  #
  # FIX (row 89): "ranged from 14 to 71" — range pattern now handles "ranged".
  # ---------------------------------------------------------------------------
  
  # --- DETECT IQR KEYWORD (used to set priority) ---
  has_iqr_keyword <- str_detect(text, regex("\\bIQR\\b|Q1.*Q3|25/?75%?\\s*IQR", ignore_case = TRUE))
  
  # --- SD ---
  sd_val <- NA_real_
  
  # FIX: "(SD) 16.8" or "(SD): 16.8" — label in brackets then number outside
  m_sd_label <- str_extract(text, regex(
    "\\(\\s*SD\\s*\\)\\s*[=:]?\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
  if (!is.na(m_sd_label)) sd_val <- as.numeric(str_extract(m_sd_label, "\\d+\\.?\\d*$"))
  
  # FIX (row 2): "mean +/- SD 75.1 ± 14.0" — only match the number AFTER ±, not before
  # "SD X" where SD is not inside brackets
  if (is.na(sd_val)) {
    m_sd1 <- str_extract(text, regex("(?<!\\()\\bSD\\b\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m_sd1)) sd_val <- as.numeric(str_extract(m_sd1, "\\d+\\.?\\d*$"))
  }
  
  # "± X" or "+/- X" or "+-X"
  # FIX (row 102): only use ± as SD when IQR keyword is NOT present
  if (is.na(sd_val) && !has_iqr_keyword) {
    m_sd2 <- str_extract(text, regex("(?:±|\\+\\s*/?\\s*-|\\+-|plus\\s*/?\\s*minus)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd2)) sd_val <- as.numeric(str_extract(m_sd2, "\\d+\\.?\\d*$"))
  }
  
  # "number±number" e.g. "62±16"
  if (is.na(sd_val) && !has_iqr_keyword) {
    m_sd3 <- str_extract(text, "\\d+\\.?\\d*\\s*±\\s*(\\d+\\.?\\d*)")
    if (!is.na(m_sd3)) sd_val <- as.numeric(str_extract(m_sd3, "\\d+\\.?\\d*$"))
  }
  
  # FIX (row 22): single bracketed number (no keyword) -> treat as SD
  # Only applies when no other dispersion type detected
  single_bracket_val <- NA_real_
  m_single <- str_extract(text, "\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
  if (!is.na(m_single)) {
    val <- as.numeric(str_extract(m_single, "\\d+\\.?\\d*"))
    # Only treat as SD if this is truly a single number (not part of a pair)
    if (!str_detect(m_single, "[,;\\-]")) single_bracket_val <- val
  }
  
  # --- IQR ---
  iqr_lo <- NA_real_
  iqr_hi <- NA_real_
  iqr_single <- NA_real_  # for cases like "IQR: 20" where only one number given
  
  m_iqr <- str_extract(text, regex(
    "(?:IQR|Q1\\s*[;,]?\\s*Q3|25/?75%?\\s*IQR)\\s*[:\\(\\[]?\\s*(\\d+\\.?\\d*)\\s*(?:[-,;]|to)\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  
  if (!is.na(m_iqr)) {
    nums <- as.numeric(str_extract_all(m_iqr, "\\d+\\.?\\d*")[[1]])
    if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
  } else if (has_iqr_keyword) {
    # Try bracketed pair when IQR keyword is elsewhere in the cell
    m_iqr2 <- str_extract(text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[,;\\-]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
    if (!is.na(m_iqr2)) {
      nums <- as.numeric(str_extract_all(m_iqr2, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
    } else {
      # FIX (row 102): IQR keyword present but only ± value given
      # Store the ± value as a single IQR dispersion measure
      m_iqr3 <- str_extract(text, regex("(?:±|\\+\\s*/?\\s*-|\\+-|\\bSD\\b)\\s*(\\d+\\.?\\d*)"))
      if (!is.na(m_iqr3)) iqr_single <- as.numeric(str_extract(m_iqr3, "\\d+\\.?\\d*$"))
    }
  }
  
  # --- RANGE ---
  # FIX (rows 86, 89, 106): expanded range pattern
  rng_lo <- NA_real_
  rng_hi <- NA_real_
  
  m_rng <- str_extract(text, regex(
    "range[d]?\\s*(?:from)?\\s*(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  if (is.na(m_rng)) {
    # Also catch "(range) X (lo to hi)" format — row 86: "median (range) age 63.8 (23-80)"
    m_rng2 <- str_extract(text, regex(
      "\\(\\s*range\\s*\\)[^\\d]*(\\d+\\.?\\d*)\\s*(?:[\\(\\[]\\s*)?(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE
    ))
    if (!is.na(m_rng2)) {
      nums <- as.numeric(str_extract_all(m_rng2, "\\d+\\.?\\d*")[[1]])
      # nums[1] = the average (e.g. 63.8), nums[2] and [3] = range bounds
      if (length(nums) >= 3) { rng_lo <- nums[2]; rng_hi <- nums[3] }
    }
  } else {
    nums <- as.numeric(str_extract_all(m_rng, "\\d+\\.?\\d*")[[1]])
    if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
  }
  
  # --- UNSPECIFIED BRACKETED PAIR ---
  unspec_lo <- NA_real_
  unspec_hi <- NA_real_
  
  # FIX (row 95): "58-62" bare hyphenated pair (no brackets, no keyword)
  # Treat as unspecified range
  m_bare_range <- str_extract(text, "^\\s*(\\d+\\.?\\d*)\\s*-\\s*(\\d+\\.?\\d*)\\s*$")
  if (!is.na(m_bare_range)) {
    nums <- as.numeric(str_extract_all(m_bare_range, "\\d+\\.?\\d*")[[1]])
    if (length(nums) == 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
  }
  
  # Bracketed pair with no keyword
  has_any_dispersion_keyword <- str_detect(text, regex(
    "\\bSD\\b|\\bIQR\\b|\\brange\\b|\\bQ1\\b|±|\\+/?-", ignore_case = TRUE))
  
  if (is.na(unspec_lo)) {
    if (!has_any_dispersion_keyword) {
      m_unspec <- str_extract(text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_unspec)) {
        nums <- as.numeric(str_extract_all(m_unspec, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
      }
    }
  }
  
  # ---------------------------------------------------------------------------
  # ASSIGN DISPERSION — priority: IQR (if keyword present) > SD > range > unspec
  # FIX (rows 81, 106): capture BOTH IQR and range when both present
  # FIX (row 102): IQR keyword takes priority over ± symbol
  # ---------------------------------------------------------------------------
  
  if (has_iqr_keyword) {
    out$measure_dispersion <- "IQR"
    if (!is.na(iqr_lo))     { out$LQR <- iqr_lo; out$UQR <- iqr_hi }
    if (!is.na(iqr_single))   out$dispersion <- iqr_single
    # Also capture range if present alongside IQR (rows 81, 106)
    if (!is.na(rng_lo))     { out$range_lower <- rng_lo; out$range_upper <- rng_hi }
    # Also capture SD if present alongside IQR
    if (!is.na(sd_val))       out$dispersion <- sd_val
    
  } else if (!is.na(sd_val)) {
    out$measure_dispersion <- "SD"
    out$dispersion         <- sd_val
    # Also capture range if present alongside SD (row 106)
    if (!is.na(rng_lo))     { out$range_lower <- rng_lo; out$range_upper <- rng_hi }
    
  } else if (!is.na(rng_lo)) {
    out$measure_dispersion <- "range"
    out$range_lower        <- rng_lo
    out$range_upper        <- rng_hi
    
  } else if (!is.na(single_bracket_val) && is.na(unspec_lo)) {
    # FIX (row 22): single bracketed number -> SD
    out$measure_dispersion <- "SD"
    out$dispersion         <- single_bracket_val
    
  } else if (!is.na(unspec_lo)) {
    out$measure_dispersion <- "unspecified"
    out$lower_unspecified  <- unspec_lo
    out$upper_unspecified  <- unspec_hi
  }
  
  # Inherit dispersion type from overall if this subgroup block has no keyword
  # but overall stated one (e.g. row 72: subgroups inherit IQR from overall)
  if (is.null(out$measure_dispersion) || is.na(out$measure_dispersion)) {
    if (!is.null(inherit_disp_type)) out$measure_dispersion <- inherit_disp_type
  }
  
  return(out)
}


# =============================================================================
# PART 2: split_groups()
# =============================================================================
# KEY FIXES vs v2:
#   - Category prefix stripped before group splitting (fixes rows 10,15,20)
#   - Value-semicolons detected and not treated as group separators (fixes 6,14)
#   - "keyword - group1 X, group2 Y" pattern added (fixes row 75)
#   - "bQ ... bM" pattern added (fixes row 90)
#   - "vs" split now cleans up leading keyword (fixes rows 50, 56)
#   - Sentence-split (row 82) improved to capture group name before first number
#   - Comma-split no longer applied to avoid row 92 false positive

split_groups <- function(text) {
  
  empty <- list(
    overall_reported   = NA_character_,
    subgroups_reported = NA_character_,
    group_1_name       = NA_character_,
    group_2_name       = NA_character_,
    overall_text       = NA_character_,
    group_1_text       = NA_character_,
    group_2_text       = NA_character_
  )
  
  if (is.na(text) || str_trim(text) == "") return(empty)
  
  text_orig <- str_trim(text)
  text <- str_replace_all(text_orig, "\u2013|\u2014", "-")
  text <- str_replace_all(text, "\u2265", ">=")
  text <- str_replace_all(text, "\u2264", "<=")
  
  # ---------------------------------------------------------------------------
  # STEP 1: Check for "overall" keyword segment
  # ---------------------------------------------------------------------------
  has_overall_keyword <- str_detect(text, regex("\\boverall\\b", ignore_case = TRUE))
  overall_text   <- NA_character_
  remaining_text <- text
  
  if (has_overall_keyword) {
    segs <- str_trim(str_split(text, ";")[[1]])
    segs <- segs[segs != ""]
    overall_idx <- which(str_detect(segs, regex("\\boverall\\b", ignore_case = TRUE)))
    if (length(overall_idx) > 0) {
      overall_text   <- segs[overall_idx[1]]
      remaining_segs <- segs[-overall_idx]
      remaining_text <- paste(remaining_segs, collapse = ";")
    }
  }
  
  # ---------------------------------------------------------------------------
  # STEP 2: Strip leading category prefix from remaining text before splitting
  # FIX (rows 10, 15, 20): prevents "≥18 years" becoming a false group_1
  # ---------------------------------------------------------------------------
  remaining_stripped <- strip_category_prefix(remaining_text)
  
  # ---------------------------------------------------------------------------
  # STEP 3: Detect group separator pattern and split
  # ---------------------------------------------------------------------------
  group_segs <- character(0)
  
  # --- Pattern A: "keyword - group1 X, group2 Y" (row 75) ---
  # e.g. "median iqr - septic shock 63 (26), noss 60 (30)"
  dash_comma_match <- str_match(remaining_stripped, regex(
    "^(?:mean|median|average|IQR|SD)?\\s*(?:\\(IQR\\)|\\(SD\\))?\\s*-\\s*(.+)",
    ignore_case = TRUE
  ))
  if (!is.na(dash_comma_match[1,1])) {
    after_dash <- str_trim(dash_comma_match[1,2])
    # Split on comma that is followed by a word (not a digit — avoids splitting numbers)
    comma_segs <- str_trim(str_split(after_dash, ",\\s*(?=[a-zA-Z])")[[1]])
    comma_segs <- comma_segs[comma_segs != ""]
    if (length(comma_segs) >= 2) {
      group_segs <- comma_segs
    }
  }
  
  # --- Pattern B: "bQ ... bM ..." (row 90) short code group labels ---
  # e.g. "bQ (n = 104) 54 (44-67) bM (n = 106) 55 (43-67)"
  if (length(group_segs) == 0) {
    bq_bm <- str_match(remaining_stripped, regex(
      "(\\b[a-z]{1,3}[Q]\\b.*?)(\\b[a-z]{1,3}[M]\\b.*)", ignore_case = TRUE))
    if (!is.na(bq_bm[1,1])) {
      group_segs <- c(str_trim(bq_bm[1,2]), str_trim(bq_bm[1,3]))
    }
  }
  
  # --- Pattern C: semicolons (most common group separator) ---
  if (length(group_segs) == 0 && str_detect(remaining_stripped, ";")) {
    # FIX (rows 6, 14): check whether semicolons are value separators first
    if (!is_value_semicolon(remaining_stripped)) {
      segs <- str_trim(str_split(remaining_stripped, ";")[[1]])
      segs <- segs[segs != ""]
      if (length(segs) >= 2) group_segs <- segs
    }
  }
  
  # --- Pattern D: "vs" separator ---
  # FIX (rows 50, 56): strip leading keyword (e.g. "Median") before using as name
  if (length(group_segs) == 0 &&
      str_detect(remaining_stripped, regex("\\bvs\\.?\\b", ignore_case = TRUE))) {
    segs <- str_trim(str_split(remaining_stripped,
                               regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  
  # --- Pattern E: sentence split on ". " before capital (for row 82, 108, 112) ---
  if (length(group_segs) == 0 && str_detect(remaining_stripped, "\\. [A-Z]")) {
    segs <- str_trim(str_split(remaining_stripped, "(?<=\\.)\\s+(?=[A-Z])")[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  
  # If no split found and no overall keyword, treat whole cell as overall
  if (length(group_segs) <= 1 && is.na(overall_text)) {
    overall_text <- text_orig
    group_segs   <- character(0)
  }
  
  # ---------------------------------------------------------------------------
  # STEP 4: Extract group name from each segment
  # FIX: for "vs" split — strip leading stat keywords like "Median" before naming
  # FIX (row 82): for sentence split — strip "Their mean ages were" type phrases
  # ---------------------------------------------------------------------------
  extract_group_name <- function(seg) {
    seg <- str_trim(seg)
    
    # Remove leading stat keyword that is NOT a group name
    # e.g. "Median 63 years (SOC)" -> strip "Median", name is "SOC"
    # e.g. "Median age: Psittacosis 48" -> strip "Median age:", name is "Psittacosis"
    seg_clean <- str_remove(seg, regex(
      "^(mean|median|average|age|years?|medican)\\s*(age\\s*)?[:\\s]*",
      ignore_case = TRUE
    ))
    
    # Remove "(n=X):" notation
    seg_clean <- str_remove(seg_clean, regex("\\(n\\s*=\\s*\\d+\\)\\s*:?\\s*", ignore_case = TRUE))
    
    # Remove "Their mean ages were" type sentence lead-ins (row 82)
    seg_clean <- str_remove(seg_clean, regex(
      "^their\\s+mean\\s+age[sd]?\\s+were?\\s*", ignore_case = TRUE))
    
    # For "vs" results: extract name from brackets if present e.g. "(SOC)"
    bracket_name <- str_extract(seg_clean, "\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(bracket_name)) {
      return(str_trim(str_remove_all(bracket_name, "[\\(\\)]")))
    }
    
    # Otherwise take text before first digit or colon
    raw_name <- str_extract(seg_clean, "^[^\\d:]+")
    if (is.na(raw_name)) return(NA_character_)
    
    # Remove trailing noise
    cleaned <- str_remove(raw_name, regex("\\s*(age\\s*)?(years?|y)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex("\\s*(through mechanical ventilation\\.?)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, ":\\s*$")
    cleaned <- str_remove(cleaned, "\\s*(Mean\\s*IQR|Mean|Median|Average)\\s*:?\\s*$")
    cleaned <- str_trim(cleaned)
    
    if (nchar(cleaned) == 0 || str_detect(cleaned, "^\\d")) return(NA_character_)
    return(cleaned)
  }
  
  group_1_name <- NA_character_
  group_2_name <- NA_character_
  group_1_text <- NA_character_
  group_2_text <- NA_character_
  
  if (length(group_segs) >= 1) {
    group_1_name <- extract_group_name(group_segs[1])
    group_1_text <- group_segs[1]
  }
  if (length(group_segs) >= 2) {
    group_2_name <- extract_group_name(group_segs[2])
    group_2_text <- group_segs[2]
  }
  
  overall_reported   <- if (!is.na(overall_text))    "Y" else "N"
  subgroups_reported <- if (length(group_segs) >= 2) "Y" else "N"
  
  return(list(
    overall_reported   = overall_reported,
    subgroups_reported = subgroups_reported,
    group_1_name       = group_1_name,
    group_2_name       = group_2_name,
    overall_text       = overall_text,
    group_1_text       = group_1_text,
    group_2_text       = group_2_text
  ))
}


# =============================================================================
# PART 3: parse_age_cell()
# =============================================================================
# FIX (row 72): overall type_of_average and measure_dispersion are passed
#   into subgroup parsing so subgroups inherit them when not restated.

parse_age_cell <- function(text) {
  
  groups <- split_groups(text)
  
  overall_data <- parse_age_block(groups$overall_text)
  
  # Pass overall's average type and dispersion type into subgroup parsing
  group_1_data <- parse_age_block(
    groups$group_1_text,
    inherit_type_avg  = overall_data$type_of_average,
    inherit_disp_type = overall_data$measure_dispersion
  )
  group_2_data <- parse_age_block(
    groups$group_2_text,
    inherit_type_avg  = overall_data$type_of_average,
    inherit_disp_type = overall_data$measure_dispersion
  )
  
  result <- c(
    list(
      overall_reported   = groups$overall_reported,
      subgroups_reported = groups$subgroups_reported,
      group_1_name       = groups$group_1_name,
      group_2_name       = groups$group_2_name
    ),
    setNames(overall_data, paste0("overall_", names(overall_data))),
    setNames(group_1_data, paste0("group_1_", names(group_1_data))),
    setNames(group_2_data, paste0("group_2_", names(group_2_data)))
  )
  
  return(result)
}


# =============================================================================
# PART 4: Apply to tibble
# =============================================================================
# Replace "my_tibble" and "age_column" with your actual names.

parsed <- lapply(age$age, parse_age_cell)

parsed_tibble <- bind_rows(
  lapply(parsed, function(x) as.data.frame(x, stringsAsFactors = FALSE))
)

my_tibble_clean <- bind_cols(age, parsed_tibble)

glimpse(my_tibble_clean)


# =============================================================================
# OPTIONAL: Save to Excel
# =============================================================================

library(openxlsx)

wb <- createWorkbook()
addWorksheet(wb, "age_v3")
writeData(wb, "age_v3", my_tibble_clean)
saveWorkbook(wb, "age_data_cleaned_v3.xlsx", overwrite = TRUE)