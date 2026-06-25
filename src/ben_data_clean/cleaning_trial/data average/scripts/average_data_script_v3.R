# =============================================================================
# Average Data Cleaning Script — Version 3
# =============================================================================
#
# FIXES vs v2
# -----------
# J  extract_bare_avg: lookahead extended to accept trailing letter so "5 IQR …"
#    extracts 5 (previously failed because "I" is not in bracket/±/+/end set)
# K  parse_average_block: inference block (after full dispersion section)
# upgrades
#    avg_not_specified to mean/median using four rules:
#    (1) lone decimal + no dispersion → mean
#    (2) unspec_disp_value present + either value has decimal → mean + SD
#    (3) IQR or LQR+UQR bounds present + integer avg → median; + decimal → mean
#    (4) SD keyword/value present → mean
#    New output flags per group: flag_inferred_mean / flag_inferred_median
# L  flag_text_only added: fires when cell has non-empty text but no digits
#    (catches "N", "NR", "not reported", etc.); included in flag_any / FLAG_COLS
# M  Inference block placed after full dispersion section so sd_val, has_sd_kw,
#    iqr_lo/iqr_hi, out$IQR_reported, out$unspec_* are all in scope
#
# FIXES vs v1
# -----------
# A  normalise_text: " | " regex bug (| unescaped = match-everything); replaced
#    with explicit unicode char class [  ·•'']
# B  normalise_text: removed no-op str_replace_all(text, "±", "±")
# C  parse_average_block: added mean Pattern 0 for "X ± Y (mean ± SD)"
# format
# D  parse_average_block: has_iqr suppressed when "no IQR" / "without IQR" in text
# E  parse_average_block: avg_not_reported uses isTRUE() to avoid NA comparison
# F  parse_average_block: post_avg position search uses word-boundary regex
#    instead of fixed() to avoid matching integer inside a larger number
# G  strip_cell_header: has_los_num check limited to text before first colon so
#    "Median (IQR) in days: no ARDS = 5 (3, 9)" is correctly detected as header;
#    data after colon is preserved in remainder
# H  extract_group_name: trailing "=" stripped from cleaned name; junk single-
#    preposition names (≤2 chars or pure stop-words) returned as NA
# I  rows_out: overall_reported / subgroups_reported removed; group_1_name /
#    group_2_name moved adjacent to their respective data blocks
#
# COLUMN MAPPING (Excel letter -> header in merged_all_sheets.xlsx)
#   D  = study number
#   F  = authors
#   I  = year published
#   R  = age
#   S  = number of participants
#   T  = groups if applicable (e.g. if case-control, cohort)
#   BB = intervention group
#   BC = control group
#   AD = icu length of stay          [PROCESSED]
#   AE = hospital length of stay     [PROCESSED]
#
# OUTPUT WORKBOOKS:
#   average_cleaned_v3.xlsx           -- full output, blank cells for NA
#   average_cleaned_v3_NR.xlsx        -- full output, "NR" for NA
#   average_cleaned_v3_condensed.xlsx -- condensed column subset
#
# =============================================================================

library(dplyr)
library(stringr)
library(openxlsx)

# =============================================================================
# CONFIGURATION
# =============================================================================

input_path <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Short-term CAP Outcomes Systematic Review (Tanvi)/CAPs_meta_repo/Meta_analysis_tanvi/data/data_extraction_tanvi_050626/merged_all_sheets.xlsx"

output_path <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Short-term CAP Outcomes Systematic Review (Tanvi)/CAPs_meta_repo/Meta_analysis_tanvi/src/ben_data_clean/cleaning_trial/average data/cleaned data/average_cleaned_v3.xlsx"

output_path_nr <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Short-term CAP Outcomes Systematic Review (Tanvi)/CAPs_meta_repo/Meta_analysis_tanvi/src/ben_data_clean/cleaning_trial/average data/cleaned data/average_cleaned_v3_NR.xlsx"

output_path_condensed <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Short-term CAP Outcomes Systematic Review (Tanvi)/CAPs_meta_repo/Meta_analysis_tanvi/src/ben_data_clean/cleaning_trial/average data/cleaned data/average_cleaned_v3_condensed.xlsx"

COL_STUDY_NUM <- "study number"
COL_AUTHORS   <- "authors"
COL_YEAR      <- "year published"
COL_GROUPS    <- "groups if applicable (e.g. if case-control, cohort)"
COL_GROUP1    <- "intervention group"
COL_GROUP2    <- "control group"
COL_AGE       <- "age"

AVERAGE_COLS <- list(
  list("AD_icu_los",      c("icu length of stay")),
  list("AE_hospital_los", c("hospital length of stay"))
)

# =============================================================================
# HELPERS
# =============================================================================

`%||%` <- function(a, b) if (!is.null(a) && length(a) > 0 && !is.na(a[1])) a else b

coalesce_val <- function(...) {
  for (a in list(...)) { if (!is.null(a) && length(a) > 0 && !is.na(a[1])) return(a) }
  NA_real_
}

# FIX A/B: removed no-op ± line; replaced " | " bug with explicit unicode class
normalise_text <- function(text) {
  if (is.na(text)) return(NA_character_)
  text <- str_replace_all(text, "–|—", "-")
  text <- str_replace_all(text, "[  ·•’‘]", " ")
  text <- str_replace_all(text, "≥", ">=")
  text <- str_replace_all(text, "≤", "<=")
  text <- str_replace_all(text, "∼", "~")
  str_trim(text)
}

is_plausible_los <- function(v) !is.na(v) && v >= 0 && v <= 9999

build_study_reference <- function(author_raw, year_raw) {
  author <- if (isTRUE(!is.na(author_raw)) && length(author_raw) > 0) str_trim(as.character(author_raw)) else ""
  year   <- if (isTRUE(!is.na(year_raw))   && length(year_raw)   > 0) str_trim(as.character(year_raw))   else ""
  if (nchar(author) == 0) {
    surname <- ""
  } else {
    first_author <- str_trim(str_split(author, regex("\\s+and\\s+", ignore_case = TRUE))[[1]][1])
    if (str_detect(first_author, ",")) {
      surname <- str_trim(str_split(first_author, ",")[[1]][1])
    } else {
      surname <- str_trim(str_split(first_author, "\\s+")[[1]][1])
    }
  }
  str_trim(paste(surname, year))
}

scalar_val <- function(x, as_type = "character") {
  if (is.null(x) || length(x) == 0)
    return(switch(as_type, character = NA_character_, numeric = NA_real_, integer = NA_integer_, NA))
  if (is.list(x)) x <- x[[1]]
  if (is.null(x) || length(x) == 0)
    return(switch(as_type, character = NA_character_, numeric = NA_real_, integer = NA_integer_, NA))
  switch(as_type,
         character = as.character(x[[1]]),
         numeric   = suppressWarnings(as.numeric(x[[1]])),
         integer   = suppressWarnings(as.integer(x[[1]])),
         x[[1]])
}

CONTRAST_PAT <- regex(
  "\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|\\bv\\b|\\band\\b|;|/|\\bgroup\\b",
  ignore_case = TRUE)

is_short_label_list <- function(g) {
  if (nchar(g) > 60) return(FALSE)
  parts <- str_trim(str_split(g, ",")[[1]])
  parts <- parts[nchar(parts) > 0]
  if (length(parts) < 2 || length(parts) > 4) return(FALSE)
  all(sapply(parts, function(p) length(str_split(p, "\\s+")[[1]]) <= 5 && !str_detect(p, "[.;]")))
}

detect_subgroups <- function(groups_val, g1_val, g2_val) {
  g  <- if (isTRUE(!is.na(groups_val)) && length(groups_val) > 0) str_trim(as.character(groups_val)) else ""
  g1 <- if (isTRUE(!is.na(g1_val))    && length(g1_val)    > 0) str_trim(as.character(g1_val))     else ""
  g2 <- if (isTRUE(!is.na(g2_val))    && length(g2_val)    > 0) str_trim(as.character(g2_val))     else ""

  if (nchar(g) > 0 && is_short_label_list(g)) {
    parts <- str_trim(str_split(g, ",")[[1]])
    parts <- parts[nchar(parts) > 0]
    return(list(has_subgroups = TRUE, n_subgroups = length(parts),
                names_from_T = parts, name_from_BB = NA_character_, name_from_BC = NA_character_))
  }
  if (str_detect(g, regex("\\bsingle\\b", ignore_case = TRUE)) && !str_detect(g, CONTRAST_PAT))
    return(list(has_subgroups = FALSE, n_subgroups = 0L, names_from_T = character(0),
                name_from_BB = NA_character_, name_from_BC = NA_character_))
  if (nchar(g1) > 0 && nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L, names_from_T = character(0),
                name_from_BB = g1, name_from_BC = g2))
  if (nchar(g1) > 0 || nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L, names_from_T = character(0),
                name_from_BB = if (nchar(g1) > 0) g1 else NA_character_,
                name_from_BC = if (nchar(g2) > 0) g2 else NA_character_))
  parts <- str_trim(str_split(g, regex("\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|;",
                                       ignore_case = TRUE))[[1]])
  parts <- parts[nchar(parts) > 0]
  if (length(parts) >= 2)
    return(list(has_subgroups = TRUE, n_subgroups = length(parts), names_from_T = parts,
                name_from_BB = NA_character_, name_from_BC = NA_character_))
  if (str_detect(g, CONTRAST_PAT))
    return(list(has_subgroups = TRUE, n_subgroups = NA_integer_, names_from_T = character(0),
                name_from_BB = NA_character_, name_from_BC = NA_character_))
  list(has_subgroups = FALSE, n_subgroups = 0L, names_from_T = character(0),
       name_from_BB = NA_character_, name_from_BC = NA_character_)
}

combine_subgroup_names <- function(names_from_T, name_from_BB, name_from_BC,
                                   cell_sub1_label, cell_sub2_label,
                                   cross_sheet_names = character(0)) {
  all_names <- c(
    if (length(names_from_T) > 0) paste0(names_from_T, " [T]") else character(0),
    if (!is.na(name_from_BB)) paste0(name_from_BB, " [BB]") else NA_character_,
    if (!is.na(name_from_BC)) paste0(name_from_BC, " [BC]") else NA_character_,
    cell_sub1_label, cell_sub2_label, cross_sheet_names
  )
  all_names <- all_names[!is.na(all_names)]
  all_names <- str_trim(all_names)
  all_names <- all_names[nchar(all_names) > 0]
  if (length(all_names) == 0) return(NA_character_)
  is_generic <- str_detect(all_names, regex("^group_\\d+(\\s*\\[[^\\]]+\\])?$", ignore_case = TRUE))
  if (any(is_generic) && any(!is_generic)) all_names <- all_names[!is_generic]
  bare <- tolower(str_remove(all_names, "\\s*\\[[^\\]]+\\]\\s*$"))
  all_names <- all_names[!duplicated(bare)]
  paste(all_names, collapse = "; ")
}

is_value_semicolon <- function(text) {
  if (!str_detect(text, ";")) return(FALSE)
  if (str_detect(text, regex("Q1\\s*;\\s*Q3", ignore_case = TRUE))) return(TRUE)
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  stat_kws <- regex(
    "^\\s*(SD|IQR|Q[123]|mean|median|medican|average|range|days?|hours?|wks?|weeks?|\\d+\\.?\\d*)\\b",
    ignore_case = TRUE)
  value_like <- sapply(segs, function(s) {
    is_stat <- str_detect(s, stat_kws)
    words   <- str_extract_all(s, "[a-zA-Z]+")[[1]]
    non_stat <- words[!str_detect(words, regex(
      "^(SD|IQR|mean|median|medican|average|range|days?|hours?|wks?|weeks?|and|with|or|the)$",
      ignore_case = TRUE))]
    has_group_name <- length(non_stat) >= 2 && str_detect(s, "\\d")
    is_stat && !has_group_name
  })
  all(value_like)
}

# FIX G: check has_los_num only in text before first colon; preserve data
# after colon in remainder so subgroup values are not lost
strip_cell_header <- function(text) {
  result <- list(header = NA_character_, remainder = text,
                 header_avg_type = NA_character_, header_disp_type = NA_character_)
  if (!str_detect(text, ";") && !str_detect(text, "\\bvs\\.?\\b") &&
      !str_detect(text, "\\.\\s+[A-Z]") && !str_detect(text, "\\s-\\s"))
    return(result)
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  if (length(segs) < 2) return(result)
  first <- segs[1]
  has_format_word <- str_detect(first, regex(
    "\\b(mean|median|medican|IQR|SD|range|average)\\b", ignore_case = TRUE))
  if (!has_format_word) return(result)

  # Only look for a LOS number in the label portion (before the first colon),
  # not in any data that follows the colon on the same segment
  first_label <- str_split(first, ":")[[1]][1]
  has_los_num <- str_detect(first_label, regex(
    "\\b\\d+(\\.\\d+)?\\s*(?:days?|hours?|wks?|[\\(\\[±])", ignore_case = TRUE))

  if (!has_los_num) {
    # If the first segment contains a colon, the data after the colon belongs
    # in the remainder so subgroup segments are not swallowed
    if (str_detect(first, ":")) {
      colon_end  <- str_locate(first, ":")[1, "end"]
      data_part  <- str_trim(substr(first, colon_end + 1, nchar(first)))
      label_part <- str_trim(substr(first, 1, colon_end))
      result$header    <- label_part
      result$remainder <- paste(c(data_part, segs[-1]), collapse = ";")
    } else {
      result$header    <- first
      result$remainder <- paste(segs[-1], collapse = ";")
    }
    result$header_avg_type  <- if (str_detect(first, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))) "median"
    else if (str_detect(first, regex("\\bmean\\b",    ignore_case = TRUE))) "mean" else NA_character_
    result$header_disp_type <- if (str_detect(first, regex("\\bIQR\\b",   ignore_case = TRUE))) "IQR"
    else if (str_detect(first, regex("\\bSD\\b",      ignore_case = TRUE))) "SD"
    else if (str_detect(first, regex("\\brange\\b",   ignore_case = TRUE))) "range" else NA_character_
  }
  result
}

# FIX J: lookahead extended to accept trailing letter so "5 IQR …" extracts 5
extract_bare_avg <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(NA_real_)
  t <- str_remove_all(text, regex("\\(\\s*n\\s*=\\s*\\d+[^)]*\\)", ignore_case = TRUE))
  t <- str_replace_all(t, regex("\\b(days?|hrs?|hours?|wks?|weeks?)\\b", ignore_case = TRUE), "")
  t <- str_trim(t)
  m <- str_extract(t, "^[^\\d]*(\\d+\\.?\\d*)\\s*(?:[\\(\\[±\\+]|$|(?=[A-Za-z]))")
  if (is.na(m)) return(NA_real_)
  val <- as.numeric(str_extract(m, "\\d+\\.?\\d*"))
  if (is_plausible_los(val)) val else NA_real_
}

# =============================================================================
# CORE: parse_average_block()
# =============================================================================
parse_average_block <- function(text,
                                inherit_mean   = NULL,
                                inherit_median = NULL,
                                inherit_disp   = NULL) {
  out <- list(
    mean_reported           = NA_character_,
    mean_value              = NA_real_,
    median_reported         = NA_character_,
    median_value            = NA_real_,
    avg_not_specified       = NA_character_,
    avg_not_specified_value = NA_real_,
    avg_not_reported        = NA_character_,
    SD_reported             = NA_character_,
    SD_value                = NA_real_,
    IQR_reported            = NA_character_,
    IQR_value               = NA_real_,
    IQR_LQR                 = NA_real_,
    IQR_UQR                 = NA_real_,
    range_reported          = NA_character_,
    range_lower             = NA_real_,
    range_upper             = NA_real_,
    unspec_reported         = NA_character_,
    unspec_lower            = NA_real_,
    unspec_upper            = NA_real_,
    unspec_disp_value       = NA_real_,
    dispersion_not_reported = NA_character_,
    flag_approx             = NA_character_,
    flag_UQR_lt_LQR         = NA_character_,
    flag_IQR_with_pm        = NA_character_,
    flag_unparsed_numbers   = NA_character_,
    flag_inferred_mean      = "N",   # FIX K
    flag_inferred_median    = "N"    # FIX K
  )

  if (is.na(text) || str_trim(text) == "") {
    out$avg_not_reported        <- "Y"
    out$dispersion_not_reported <- "Y"
    out[c("mean_reported","median_reported","avg_not_specified",
          "SD_reported","IQR_reported","range_reported","unspec_reported")] <- "N"
    out[c("flag_approx","flag_UQR_lt_LQR","flag_IQR_with_pm","flag_unparsed_numbers")] <- "N"
    return(out)
  }

  text <- normalise_text(text)
  out$flag_approx      <- if (str_detect(text, "~")) "Y" else "N"
  out$flag_IQR_with_pm <- if (str_detect(text, regex("\\bIQR\\b", ignore_case = TRUE)) &&
                              str_detect(text, "[±]|\\+/?-|\\+-")) "Y" else "N"

  tc <- str_replace_all(text, regex("\\b(days?|hrs?|hours?|wks?|weeks?)\\b", ignore_case = TRUE), "")
  tc <- str_trim(str_replace_all(tc, "\\s+", " "))

  # -----------------------------------------------------------------------
  # AVERAGE EXTRACTION
  # -----------------------------------------------------------------------

  median_secondary <- NA_real_
  sec_m <- str_extract(tc, regex("\\(\\s*median\\s*,?\\s*(\\d+\\.?\\d*)\\s*\\)", ignore_case = TRUE))
  if (!is.na(sec_m)) median_secondary <- as.numeric(str_extract(sec_m, "\\d+\\.?\\d*$"))
  tc_no_sec <- str_remove(tc, regex("\\(\\s*median\\s*,?\\s*\\d+\\.?\\d*\\s*\\)", ignore_case = TRUE))

  # --- Mean ---
  mean_val    <- NA_real_
  has_mean_kw <- str_detect(tc_no_sec, regex("\\bmean\\b", ignore_case = TRUE))

  if (has_mean_kw) {
    # "mean +/- SD X ± Y" -> X is mean
    m <- str_extract(tc_no_sec, regex("\\bmean\\s*\\+/?-\\s*SD\\s+(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m)) mean_val <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))

    # FIX C: Pattern 0 — "X ± Y (mean ± SD)" or "X ± Y (mean)" where first number is mean
    if (is.na(mean_val)) {
      m <- str_extract(tc_no_sec, regex(
        "(\\d+\\.?\\d*)\\s*(?:[±]|\\+/?-)\\s*\\d+\\.?\\d*\\s*\\(\\s*mean", ignore_case = TRUE))
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "^\\d+\\.?\\d*")); if (is_plausible_los(v)) mean_val <- v }
    }
    # Pattern A: keyword + optional words + number
    if (is.na(mean_val)) {
      m <- str_extract(tc_no_sec, regex(
        "\\bmean\\s*(?:length\\s*of\\s*stay|los|icu\\s*(?:stay|los)?)?\\s*(?:(?:was|were|of|at)\\s*)?[~=:]?\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*$")); if (is_plausible_los(v)) mean_val <- v }
    }
    # Pattern B: "Mean 7.5"
    if (is.na(mean_val)) {
      m <- str_extract(tc_no_sec, regex("\\bmean\\s+(\\d+\\.?\\d*)\\b", ignore_case = TRUE))
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*")); if (is_plausible_los(v)) mean_val <- v }
    }
    # Pattern C: "7.5 (mean)"
    if (is.na(mean_val)) {
      m <- str_extract(tc_no_sec, regex("(\\d+\\.?\\d*)\\s*\\(\\s*mean\\s*\\)", ignore_case = TRUE))
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "^\\d+\\.?\\d*")); if (is_plausible_los(v)) mean_val <- v }
    }
    # Pattern D: "Mean (SD) X"
    if (is.na(mean_val)) {
      m <- str_extract(tc_no_sec, regex(
        "\\bmean\\s*\\([^)]*\\)\\s*[=:]?\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*$")); if (is_plausible_los(v)) mean_val <- v }
    }
    # Pattern E: fallback bare number after keyword
    if (is.na(mean_val)) {
      stripped <- str_remove(tc_no_sec, regex(
        ".*?\\bmean\\b\\s*(?:length\\s*of\\s*stay|los)?\\s*[=:,]?\\s*", ignore_case = TRUE))
      stripped <- str_remove(stripped, regex("^\\s*\\([^\\d)]*\\)\\s*", ignore_case = TRUE))
      stripped <- str_remove_all(stripped, regex("\\(\\s*n\\s*=\\s*\\d+[^)]*\\)", ignore_case = TRUE))
      m <- str_extract(stripped, "^[^\\d]*(\\d+\\.?\\d*)")
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*")); if (is_plausible_los(v)) mean_val <- v }
    }
  }

  # --- Median ---
  median_val    <- NA_real_
  has_median_kw <- str_detect(tc, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))

  if (has_median_kw) {
    tc_no_n <- str_remove_all(tc, regex("\\(\\s*n\\s*=\\s*\\d+[^)]*\\)", ignore_case = TRUE))
    m <- str_extract(tc_no_n, regex(
      "\\b(?:median|medican)\\b[^0-9]*?(\\b\\d{1,5}(?:\\.\\d+)?\\b)", ignore_case = TRUE))
    if (!is.na(m)) { val <- as.numeric(str_extract(m, "[\\d.]+$")); if (is_plausible_los(val)) median_val <- val }

    if (is.na(median_val)) {
      m <- str_extract(tc, regex("(\\d+\\.?\\d*)\\s*(?:\\([^)]*\\))?\\s*(?:median|medican)", ignore_case = TRUE))
      if (!is.na(m)) { val <- as.numeric(str_extract(m, "^\\d+\\.?\\d*")); if (is_plausible_los(val)) median_val <- val }
    }
    if (is.na(median_val)) {
      if (str_detect(tc, regex("\\(\\s*(?:median|medican)\\s*[±]", ignore_case = TRUE))) {
        first_num <- str_extract(tc, "\\d+\\.?\\d*")
        if (!is.na(first_num)) { val <- as.numeric(first_num); if (is_plausible_los(val)) median_val <- val }
      }
    }
    if (is.na(median_val)) {
      m <- str_extract(tc, regex("\\bmedian\\b\\s*~\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m)) { val <- as.numeric(str_extract(m, "\\d+\\.?\\d*$")); if (is_plausible_los(val)) median_val <- val }
    }
  }

  if (!is.na(median_secondary) && is.na(median_val)) median_val <- median_secondary

  # --- Average not specified ---
  avg_ns_val <- NA_real_
  has_avg_kw <- str_detect(tc, regex("\\baverage\\b", ignore_case = TRUE))

  if (!has_mean_kw && !has_median_kw) {
    avg_ns_val <- extract_bare_avg(tc)
  } else if (has_avg_kw && !has_mean_kw) {
    m <- str_extract(tc, regex(
      "\\baverage\\s*(?:length\\s*of\\s*stay|los)?\\s*[=:]?\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*$")); if (is_plausible_los(v)) avg_ns_val <- v }
  }

  # FIX E: use isTRUE() to guard NA comparisons
  out$mean_reported           <- if (!is.na(mean_val)    || has_mean_kw)   "Y" else "N"
  out$mean_value              <- mean_val
  out$median_reported         <- if (!is.na(median_val)  || has_median_kw) "Y" else "N"
  out$median_value            <- median_val
  out$avg_not_specified       <- if (!is.na(avg_ns_val)  || has_avg_kw)    "Y" else "N"
  out$avg_not_specified_value <- avg_ns_val
  out$avg_not_reported        <- if (isTRUE(out$mean_reported     == "N") &&
                                      isTRUE(out$median_reported   == "N") &&
                                      isTRUE(out$avg_not_specified == "N")) "Y" else "N"

  # Inheritance: propagate avg type from overall/cell header to subgroups
  if (isTRUE(out$avg_not_reported == "Y") || isTRUE(out$avg_not_specified == "Y")) {
    bare_val <- if (isTRUE(out$avg_not_reported == "Y")) extract_bare_avg(tc) else avg_ns_val
    if (!is.null(inherit_median) && isTRUE(inherit_median == "Y") &&
        isTRUE(out$median_reported == "N")) {
      out$median_reported  <- "Y"
      out$median_value     <- bare_val
      out$avg_not_reported <- "N"
      if (isTRUE(out$avg_not_specified == "Y") && identical(out$avg_not_specified_value, bare_val)) {
        out$avg_not_specified <- "N"; out$avg_not_specified_value <- NA_real_
      }
      out$median_value <- bare_val
    } else if (!is.null(inherit_mean) && isTRUE(inherit_mean == "Y") &&
               isTRUE(out$mean_reported == "N")) {
      out$mean_reported    <- "Y"
      out$mean_value       <- bare_val
      out$avg_not_reported <- "N"
      if (isTRUE(out$avg_not_specified == "Y") && identical(out$avg_not_specified_value, bare_val)) {
        out$avg_not_specified <- "N"; out$avg_not_specified_value <- NA_real_
      }
      out$mean_value <- bare_val
    }
  }

  # FIX F: use word-boundary regex to locate avg in text (avoids integer substring match)
  avg_used <- coalesce_val(mean_val, median_val, avg_ns_val,
                           out$mean_value, out$median_value, out$avg_not_specified_value)
  post_avg <- tc
  if (!is.na(avg_used)) {
    avg_str <- as.character(avg_used)
    avg_pat <- paste0("(?<![.\\d])", gsub("\\.", "\\\\.", avg_str), "(?![.\\d])")
    m_pos   <- str_locate(tc, regex(avg_pat))
    pos     <- m_pos[1, "end"]
    if (!is.na(pos)) post_avg <- substr(tc, pos + 1, nchar(tc))
  }

  # -----------------------------------------------------------------------
  # DISPERSION EXTRACTION
  # -----------------------------------------------------------------------
  has_iqr   <- str_detect(tc, regex("\\bIQR\\b|Q1.*Q3|25/?75%?\\s*IQR", ignore_case = TRUE))
  # FIX D: suppress IQR flag when the text explicitly says "no IQR"
  if (has_iqr && str_detect(tc, regex(
    "\\bno\\s+IQR\\b|\\bwithout\\s+IQR\\b|\\bnot.*\\bIQR\\b", ignore_case = TRUE)))
    has_iqr <- FALSE

  has_range <- str_detect(tc, regex("\\brange[d]?\\b", ignore_case = TRUE))
  has_sd_kw <- str_detect(tc, regex("\\bSD\\b|[±]|\\+\\s*/?\\s*-|\\+-"))

  # --- SD ---
  sd_val <- NA_real_
  m <- str_extract(tc, regex(
    "\\bmean\\s*\\+/?-\\s*SD\\s+\\d+\\.?\\d*\\s*(?:[±]|\\+/?-)\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
  if (!is.na(m)) { nums <- as.numeric(str_extract_all(m, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) sd_val <- nums[length(nums)] }

  if (is.na(sd_val) && has_sd_kw) {
    m <- str_extract(post_avg, regex(
      "\\(\\s*SD\\s*\\)\\s*[=:]?\\s*(\\d+\\.?\\d*)|\\(\\s*SD\\s+(\\d+\\.?\\d*)\\s*\\)", ignore_case = TRUE))
    if (!is.na(m)) { nums <- str_extract_all(m, "\\d+\\.?\\d*")[[1]]; if (length(nums) > 0) sd_val <- as.numeric(nums[length(nums)]) }
    if (is.na(sd_val) && !has_iqr) {
      m <- str_extract(post_avg, "\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*")); if (!is.na(v) && !str_detect(m, "[,;\\-]")) sd_val <- v }
    }
  }
  if (is.na(sd_val) && !has_iqr) {
    m <- str_extract(post_avg, regex("\\bSD\\b\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m)) sd_val <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))
  }
  if (is.na(sd_val) && !has_iqr) {
    m <- str_extract(post_avg, regex("(?:[±]|\\+\\s*/?\\s*-|\\+-|plus\\s*/?\\s*minus)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m)) sd_val <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))
  }
  if (is.na(sd_val) && !has_iqr) {
    m <- str_extract(post_avg, regex("(?<![+/\\-])\\+(?![+/\\-])\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m)) sd_val <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))
  }
  if (is.na(sd_val) && str_detect(tc, regex("\\(\\s*(?:median|medican)\\s*[±]", ignore_case = TRUE))) {
    m <- str_extract(tc, regex("(\\d+\\.?\\d*)\\s*\\(\\s*(?:median|medican)\\s*[±]", ignore_case = TRUE))
    if (!is.na(m)) sd_val <- as.numeric(str_extract(m, "^\\d+\\.?\\d*"))
  }

  # --- IQR ---
  iqr_val <- NA_real_; iqr_lo <- NA_real_; iqr_hi <- NA_real_
  if (has_iqr) {
    m <- str_extract(tc, regex(
      "(?:IQR|Q1\\s*[;,]?\\s*Q3|25/?75%?\\s*IQR)\\s*[:\\(\\[]?\\s*(\\d+\\.?\\d*)\\s*(?:[-,;]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m)) { nums <- as.numeric(str_extract_all(m, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] } }
    if (is.na(iqr_lo)) {
      m <- str_extract(post_avg, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[,;\\-]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m)) { nums <- as.numeric(str_extract_all(m, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] } }
    }
    if (is.na(iqr_lo)) {
      m <- str_extract(post_avg, regex(
        "IQR\\s*:?\\s*(\\d+\\.?\\d*)|(?:[±]|\\+\\s*/?\\s*-|\\+-)\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m)) { nums <- as.numeric(str_extract_all(m, "\\d+\\.?\\d*")[[1]]); if (length(nums) > 0) iqr_val <- nums[length(nums)] }
    }
    if (is.na(iqr_lo) && is.na(iqr_val) && !is.null(inherit_disp) && isTRUE(inherit_disp == "IQR")) {
      m <- str_extract(post_avg, "\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
      if (!is.na(m)) { v <- as.numeric(str_extract(m, "\\d+\\.?\\d*")); if (!is.na(v) && !str_detect(m, "[,;\\-]")) iqr_val <- v }
    }
  }

  # --- Range ---
  rng_lo <- NA_real_; rng_hi <- NA_real_
  if (has_range) {
    m <- str_extract(tc, regex("range[d]?\\s*(?:from)?\\s*(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m)) {
      nums <- as.numeric(str_extract_all(m, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
    } else {
      m2 <- str_extract(tc, regex(
        "\\(\\s*range\\s*\\)[^\\d]*(\\d+\\.?\\d*)[^\\d]+(\\d+\\.?\\d*)\\s*(?:to|-|,)\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m2)) {
        nums <- as.numeric(str_extract_all(m2, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 3) { rng_lo <- nums[2]; rng_hi <- nums[3] }
      } else {
        m3 <- str_extract(post_avg, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
        if (!is.na(m3)) { nums <- as.numeric(str_extract_all(m3, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] } }
      }
    }
  }

  # --- Unspecified dispersion ---
  unspec_lo <- NA_real_; unspec_hi <- NA_real_; unspec_val <- NA_real_

  m_bare <- str_extract(str_trim(tc), "^\\s*(\\d+\\.?\\d*)\\s*-\\s*(\\d+\\.?\\d*)\\s*$")
  if (!is.na(m_bare)) {
    nums <- as.numeric(str_extract_all(m_bare, "\\d+\\.?\\d*")[[1]])
    if (length(nums) == 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
  }

  has_any_kw <- str_detect(tc, regex("\\bSD\\b|\\bIQR\\b|\\brange\\b|\\bQ1\\b|[±]|\\+/?-", ignore_case = TRUE))
  if (is.na(unspec_lo) && !has_any_kw) {
    m_s <- str_extract(post_avg, "^\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
    if (!is.na(m_s) && !str_detect(m_s, "[,;\\-]")) {
      unspec_val <- as.numeric(str_extract(m_s, "\\d+\\.?\\d*"))
    } else {
      m_p <- str_extract(post_avg, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_p)) { nums <- as.numeric(str_extract_all(m_p, "\\d+\\.?\\d*")[[1]]); if (length(nums) >= 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] } }
    }
  }

  # -----------------------------------------------------------------------
  # ASSIGN DISPERSION FLAGS
  # -----------------------------------------------------------------------
  any_disp <- FALSE

  if (!is.na(sd_val)) {
    out$SD_reported <- "Y"; out$SD_value <- sd_val; any_disp <- TRUE
  } else { out$SD_reported <- "N" }

  if (has_iqr) {
    out$IQR_reported <- "Y"; any_disp <- TRUE
    out$IQR_LQR <- iqr_lo; out$IQR_UQR <- iqr_hi; out$IQR_value <- iqr_val
    if (!is.na(rng_lo)) { out$range_reported <- "Y"; out$range_lower <- rng_lo; out$range_upper <- rng_hi }
  } else { out$IQR_reported <- "N" }

  if (!is.na(rng_lo) && !isTRUE(out$range_reported == "Y")) {
    out$range_reported <- "Y"; out$range_lower <- rng_lo; out$range_upper <- rng_hi; any_disp <- TRUE
  } else if (!isTRUE(out$range_reported == "Y")) { out$range_reported <- "N" }

  if (!is.na(unspec_lo) || !is.na(unspec_val)) {
    out$unspec_reported   <- "Y"
    out$unspec_lower      <- unspec_lo
    out$unspec_upper      <- unspec_hi
    out$unspec_disp_value <- unspec_val
    any_disp <- TRUE
  } else { out$unspec_reported <- "N" }

  out$dispersion_not_reported <- if (any_disp) "N" else "Y"

  if (isTRUE(out$dispersion_not_reported == "Y") && !is.null(inherit_disp)) {
    if (isTRUE(inherit_disp == "IQR"))   { out$IQR_reported   <- "Y"; out$dispersion_not_reported <- "N" }
    if (isTRUE(inherit_disp == "SD"))    { out$SD_reported    <- "Y"; out$dispersion_not_reported <- "N" }
    if (isTRUE(inherit_disp == "range")) { out$range_reported <- "Y"; out$dispersion_not_reported <- "N" }
  }
  if (!is.null(inherit_disp) && isTRUE(inherit_disp == "IQR") && isTRUE(out$unspec_reported == "Y")) {
    out$IQR_reported <- "Y"; out$dispersion_not_reported <- "N"
    if (!is.na(out$unspec_lower))      { out$IQR_LQR <- out$unspec_lower; out$IQR_UQR <- out$unspec_upper; out$unspec_lower <- NA_real_; out$unspec_upper <- NA_real_ }
    if (!is.na(out$unspec_disp_value)) { out$IQR_value <- out$unspec_disp_value; out$unspec_disp_value <- NA_real_ }
    out$unspec_reported <- "N"
  }

  # -----------------------------------------------------------------------
  # FIX K: INFERENCE — upgrade avg_not_specified to mean/median
  # Runs after full dispersion section so sd_val, has_sd_kw, iqr_lo/hi,
  # out$IQR_reported, out$unspec_* are all in scope (FIX M)
  # -----------------------------------------------------------------------
  if (isTRUE(out$avg_not_specified == "Y") && !is.na(out$avg_not_specified_value)) {
    val        <- out$avg_not_specified_value
    has_dec    <- grepl("\\.", as.character(val))
    unspec_dec <- !is.na(out$unspec_disp_value) && grepl("\\.", as.character(out$unspec_disp_value))

    # has_bounds: explicit IQR keyword, extracted IQR values, IQR inherited,
    # or a lower+upper unspecified boundary pair
    has_bounds <- has_iqr ||
                  (!is.na(iqr_lo) && !is.na(iqr_hi)) ||
                  isTRUE(out$IQR_reported == "Y") ||
                  (!is.na(out$unspec_lower) && !is.na(out$unspec_upper))

    inf <- NA_character_
    if (!is.na(sd_val) || (has_sd_kw && !has_iqr)) {
      inf <- "mean"                                         # Assumption 4: SD → mean
    } else if (!is.na(out$unspec_disp_value) && (has_dec || unspec_dec)) {
      inf <- "mean_sd"                                      # Assumption 2: single bracket + decimal → mean+SD
    } else if (has_bounds) {
      inf <- if (has_dec) "mean" else "median"              # Assumption 3: IQR/bounds present
    } else if (has_dec && isTRUE(out$dispersion_not_reported == "Y")) {
      inf <- "mean"                                         # Assumption 1: lone decimal, no dispersion → mean
    }

    if (!is.na(inf)) {
      if (inf %in% c("mean", "mean_sd")) {
        out$mean_reported      <- "Y"
        out$mean_value         <- val
        out$flag_inferred_mean <- "Y"
        if (inf == "mean_sd" && !is.na(out$unspec_disp_value)) {
          out$SD_reported             <- "Y"
          out$SD_value                <- out$unspec_disp_value
          out$unspec_reported         <- "N"
          out$unspec_disp_value       <- NA_real_
          out$dispersion_not_reported <- "N"
        }
      } else {
        out$median_reported        <- "Y"
        out$median_value           <- val
        out$flag_inferred_median   <- "Y"
      }
      out$avg_not_specified       <- "N"
      out$avg_not_specified_value <- NA_real_
      out$avg_not_reported        <- "N"
    }
  }

  out$flag_UQR_lt_LQR <- if (!is.na(iqr_lo) && !is.na(iqr_hi) && iqr_hi < iqr_lo) "Y" else "N"

  all_raw  <- as.numeric(str_extract_all(tc, "\\d+\\.?\\d*")[[1]])
  all_raw  <- all_raw[!is.na(all_raw)]
  out_nums <- unlist(out[sapply(out, is.numeric)])
  out_nums <- as.numeric(out_nums[!is.na(out_nums)])
  out$flag_unparsed_numbers <- if (length(all_raw) > (length(out_nums) + 1)) "Y" else "N"

  out
}

# =============================================================================
# split_groups()
# =============================================================================
split_groups <- function(text) {
  empty <- list(
    overall_reported   = NA_character_,
    subgroups_reported = NA_character_,
    group_1_name       = NA_character_,
    group_2_name       = NA_character_,
    overall_text       = NA_character_,
    group_1_text       = NA_character_,
    group_2_text       = NA_character_,
    cell_avg_type      = NA_character_,
    cell_disp_type     = NA_character_
  )
  if (is.na(text) || str_trim(text) == "") return(empty)

  text_orig <- normalise_text(text)
  text_work <- str_remove_all(text_orig, '^"|"$')
  text_work <- normalise_text(text_work)

  header_info  <- strip_cell_header(text_work)
  working_text <- normalise_text(header_info$remainder)

  cell_avg_type  <- header_info$header_avg_type
  cell_disp_type <- header_info$header_disp_type
  if (is.na(cell_avg_type))
    cell_avg_type  <- if (str_detect(text_work, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))) "median"
  else if (str_detect(text_work, regex("\\bmean\\b", ignore_case = TRUE))) "mean" else NA_character_
  if (is.na(cell_disp_type))
    cell_disp_type <- if (str_detect(text_work, regex("\\bIQR\\b",   ignore_case = TRUE))) "IQR"
  else if (str_detect(text_work, regex("\\bSD\\b",    ignore_case = TRUE))) "SD"
  else if (str_detect(text_work, regex("\\brange\\b", ignore_case = TRUE))) "range" else NA_character_

  has_overall_kw <- str_detect(working_text, regex("\\boverall\\b", ignore_case = TRUE))
  overall_text   <- NA_character_
  remaining_text <- working_text

  if (has_overall_kw) {
    segs   <- str_trim(str_split(working_text, ";")[[1]])
    segs   <- segs[segs != ""]
    ov_idx <- which(str_detect(segs, regex("\\boverall\\b", ignore_case = TRUE)))
    if (length(ov_idx) > 0) {
      overall_text   <- segs[ov_idx[1]]
      remaining_text <- paste(segs[-ov_idx], collapse = ";")
    }
  }

  group_segs <- character(0)

  if (length(group_segs) == 0) {
    bq_bm <- str_match(remaining_text, regex("(\\b[a-z]{1,3}Q\\b.+?)(\\b[a-z]{1,3}M\\b.+)", ignore_case = TRUE))
    if (!is.na(bq_bm[1, 1])) group_segs <- c(str_trim(bq_bm[1, 2]), str_trim(bq_bm[1, 3]))
  }
  if (length(group_segs) == 0 &&
      str_count(remaining_text, regex("\\(n\\s*=\\s*\\d+\\)")) >= 2) {
    parts <- str_trim(unlist(str_split(remaining_text, regex("(?<=\\d\\)\\s{0,3})(?=[a-zA-Z])"))))
    parts <- parts[parts != "" & str_detect(parts, "\\d")]
    if (length(parts) >= 2) group_segs <- parts
  }
  if (length(group_segs) == 0 && str_detect(remaining_text, ";")) {
    if (!is_value_semicolon(remaining_text)) {
      segs <- str_trim(str_split(remaining_text, ";")[[1]])
      segs <- segs[segs != ""]
      if (length(segs) >= 2) group_segs <- segs
    }
  }
  if (length(group_segs) == 0 &&
      str_detect(remaining_text, regex("\\bvs\\.?\\b", ignore_case = TRUE))) {
    segs <- str_trim(str_split(remaining_text, regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  if (length(group_segs) == 0 && str_detect(remaining_text, "\\.\\s+\\w")) {
    segs <- str_trim(str_split(remaining_text, "(?<=\\.)\\s+(?=\\w)")[[1]])
    segs <- segs[segs != ""]
    pronouns <- regex("^(Their|His|Her|Its|The|A|An|This|These|Those)\\b", ignore_case = FALSE)
    merged <- character(0)
    for (s in segs) {
      if (length(merged) > 0 && str_detect(s, pronouns)) merged[length(merged)] <- paste(merged[length(merged)], s)
      else merged <- c(merged, s)
    }
    if (length(merged) >= 2) group_segs <- merged
  }

  if (length(group_segs) <= 1 && is.na(overall_text)) {
    overall_text <- text_orig
    group_segs   <- character(0)
  }

  # FIX H: extract_group_name now strips trailing "=" and rejects junk words
  extract_group_name <- function(seg) {
    seg       <- str_trim(seg)
    seg_clean <- str_remove(seg, regex("\\(n\\s*=\\s*\\d+[^)]*\\)\\s*:?\\s*", ignore_case = TRUE))
    # Bracketed name at end
    end_brk <- str_extract(seg_clean, regex("\\(\\s*([A-Z][A-Za-z0-9\\-\\s]+)\\s*\\)\\s*$"))
    if (!is.na(end_brk)) {
      nm <- str_trim(str_remove_all(end_brk, "[\\(\\)]"))
      if (nchar(nm) >= 2 && !str_detect(nm, regex(
        "^(IQR|SD|range|mean|median|average|days?|hours?)$", ignore_case = TRUE))) return(nm)
    }
    # Plain words at end after numbers/brackets
    end_words <- str_extract(seg_clean, regex("(?:[\\d\\s\\(\\)\\[\\].,;\\u00b1+-]+)([a-zA-Z][a-zA-Z\\s-]+)$"))
    if (!is.na(end_words)) {
      nm <- str_trim(str_extract(end_words, "[a-zA-Z][a-zA-Z\\s-]+$"))
      if (!is.na(nm) && nchar(nm) >= 3 &&
          !str_detect(nm, regex("^(days?|hours?|wks?|weeks?|SD|IQR|range|mean|median|average)$", ignore_case = TRUE)))
        return(str_trim(nm))
    }
    seg_clean <- str_remove(seg_clean, regex("^(mean|median|medican|average|IQR|SD)\\s*[:\\s]*", ignore_case = TRUE))
    brk_start <- str_extract(seg_clean, "^\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(brk_start)) return(str_trim(str_remove_all(brk_start, "[\\(\\)]")))
    brk_mid <- str_extract(seg_clean, "\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(brk_mid)) {
      nm <- str_trim(str_remove_all(brk_mid, "[\\(\\)]"))
      if (nchar(nm) >= 2 && !str_detect(nm, regex("^(IQR|SD|range|mean|median)$", ignore_case = TRUE))) return(nm)
    }
    raw_nm <- str_extract(seg_clean, "^[^\\d:]+")
    if (is.na(raw_nm)) return(NA_character_)
    cleaned <- str_remove(raw_nm, regex("\\s*(days?|hours?|wks?|weeks?)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex("\\s*\\(\\s*(range|IQR|SD|mean|median)\\s*\\)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex("\\s*(mean|median|IQR|SD|average)\\s*:?\\s*$", ignore_case = TRUE))
    cleaned <- str_trim(str_remove(cleaned, "[=:]\\s*$"))   # strip trailing = or :
    cleaned <- str_trim(cleaned)
    if (nchar(cleaned) == 0 || str_detect(cleaned, "^[\\d>=<]")) return(NA_character_)
    # Reject pure stop-words or very short fragments
    JUNK_PAT <- regex("^(in|of|at|no|and|or|the|with|by|from|to|for|a|an)$", ignore_case = TRUE)
    if (str_detect(cleaned, JUNK_PAT) || nchar(cleaned) < 3) return(NA_character_)
    # Reject if starts with a stat keyword (e.g. "days", "median")
    if (str_detect(cleaned, regex(
      "^(median|mean|IQR|SD|range|average|days?|hours?)\\b", ignore_case = TRUE)))
      return(NA_character_)
    cleaned
  }

  g1_name <- NA_character_; g2_name <- NA_character_
  g1_text <- NA_character_; g2_text <- NA_character_
  if (length(group_segs) >= 1) { g1_name <- extract_group_name(group_segs[1]); g1_text <- group_segs[1] }
  if (length(group_segs) >= 2) { g2_name <- extract_group_name(group_segs[2]); g2_text <- group_segs[2] }

  list(
    overall_reported   = if (!is.na(overall_text))    "Y" else "N",
    subgroups_reported = if (length(group_segs) >= 2) "Y" else "N",
    group_1_name       = g1_name,
    group_2_name       = g2_name,
    overall_text       = overall_text,
    group_1_text       = g1_text,
    group_2_text       = g2_text,
    cell_avg_type      = cell_avg_type,
    cell_disp_type     = cell_disp_type
  )
}

# =============================================================================
# parse_average_cell(): top-level function for one cell
# =============================================================================
parse_average_cell <- function(text) {
  groups       <- split_groups(text)
  overall_data <- parse_average_block(groups$overall_text)

  inh_mean   <- if (isTRUE(overall_data$mean_reported   == "Y")) "Y"
  else if (!is.na(groups$cell_avg_type) && groups$cell_avg_type == "mean")   "Y" else "N"
  inh_median <- if (isTRUE(overall_data$median_reported == "Y")) "Y"
  else if (!is.na(groups$cell_avg_type) && groups$cell_avg_type == "median") "Y" else "N"
  inh_disp   <- if (isTRUE(overall_data$IQR_reported == "Y"))  "IQR"
  else if (isTRUE(overall_data$SD_reported  == "Y"))  "SD"
  else if (!is.na(groups$cell_disp_type))              groups$cell_disp_type
  else NULL

  g1_data <- parse_average_block(groups$group_1_text,
                                 inherit_mean = inh_mean, inherit_median = inh_median, inherit_disp = inh_disp)
  g2_data <- parse_average_block(groups$group_2_text,
                                 inherit_mean = inh_mean, inherit_median = inh_median, inherit_disp = inh_disp)

  c(
    list(overall_reported   = groups$overall_reported,
         subgroups_reported = groups$subgroups_reported,
         group_1_name       = groups$group_1_name,
         group_2_name       = groups$group_2_name),
    setNames(overall_data, paste0("overall_",  names(overall_data))),
    setNames(g1_data,      paste0("group_1_",  names(g1_data))),
    setNames(g2_data,      paste0("group_2_",  names(g2_data)))
  )
}

# =============================================================================
# PART 1: Read data
# =============================================================================
df <- read.xlsx(input_path, sheet = 1, check.names = FALSE, sep.names = " ")

df[] <- lapply(df, function(col) {
  if (is.numeric(col)) {
    vapply(col, function(x) if (is.na(x)) NA_character_ else as.character(signif(x, 10)), character(1))
  } else {
    as.character(col)
  }
})

expected_cols <- c(COL_STUDY_NUM, COL_AUTHORS, COL_YEAR, COL_GROUPS, COL_GROUP1, COL_GROUP2, COL_AGE,
                   unlist(lapply(AVERAGE_COLS, function(x) x[[2]])))
missing_cols  <- setdiff(expected_cols, colnames(df))
if (length(missing_cols) > 0)
  warning("The following expected columns were NOT found in the data: ",
          paste(missing_cols, collapse = ", "))

# =============================================================================
# PART 2: Cross-sheet subgroup name pre-scan
# =============================================================================
cell_level_names_by_study <- new.env(hash = TRUE)

for (col_def in AVERAGE_COLS) {
  existing_cols <- intersect(col_def[[2]], colnames(df))
  if (length(existing_cols) == 0) next
  tag <- existing_cols[1]

  for (i in seq_len(nrow(df))) {
    study_key <- scalar_val(df[i, ][[COL_STUDY_NUM]], "character")
    if (is.na(study_key) || nchar(study_key) == 0) next
    cell_val  <- scalar_val(df[i, existing_cols[1]], "character")
    parsed    <- parse_average_cell(cell_val)
    found <- c(parsed$group_1_name, parsed$group_2_name)
    found <- found[!is.na(found) & !str_detect(found, regex("^group_\\d+$", ignore_case = TRUE))]
    if (length(found) == 0) next
    tagged   <- paste0(found, " [", tag, "]")
    existing <- if (exists(study_key, envir = cell_level_names_by_study, inherits = FALSE))
      get(study_key, envir = cell_level_names_by_study) else character(0)
    assign(study_key, c(existing, tagged), envir = cell_level_names_by_study)
  }
}

get_cell_level_names_for_study <- function(study_key) {
  if (is.na(study_key) ||
      !exists(study_key, envir = cell_level_names_by_study, inherits = FALSE)) return(character(0))
  get(study_key, envir = cell_level_names_by_study)
}

# =============================================================================
# PART 3: Process each average column
# =============================================================================
all_sheet_data <- list()

for (col_def in AVERAGE_COLS) {
  sheet_name    <- col_def[[1]]
  src_col_names <- col_def[[2]]
  cat("Processing", sheet_name, "...\n")

  existing_cols <- intersect(src_col_names, colnames(df))
  if (length(existing_cols) == 0) {
    message("  No source columns found for ", sheet_name, " -- skipping.")
    next
  }
  current_tag <- existing_cols[1]

  rows_out <- vector("list", nrow(df))

  for (i in seq_len(nrow(df))) {
    row      <- df[i, ]
    cell_val <- scalar_val(row[[existing_cols[1]]], "character")

    study_num <- scalar_val(row[[COL_STUDY_NUM]], "character")
    study_ref <- build_study_reference(scalar_val(row[[COL_AUTHORS]], "character"),
                                       scalar_val(row[[COL_YEAR]],    "character"))

    sub_info <- detect_subgroups(
      scalar_val(row[[COL_GROUPS]], "character"),
      scalar_val(row[[COL_GROUP1]], "character"),
      scalar_val(row[[COL_GROUP2]], "character"))

    parsed <- parse_average_cell(cell_val)

    flag_not      <- if (!is.na(cell_val) && str_detect(cell_val, regex("\\bnot\\b", ignore_case = TRUE))) "Y" else "N"
    flag_question <- if (!is.na(cell_val) && str_detect(cell_val, "\\?")) "Y" else "N"

    flag_prop_data <- if (!is.na(cell_val) &&
                          (str_detect(cell_val, "\\d+\\s*/\\s*\\d+") ||
                           str_detect(cell_val, "%") ||
                           str_detect(cell_val, "\\b0\\.\\d+"))) "Y" else "N"

    all_raw_nums <- if (!is.na(cell_val))
      as.numeric(str_extract_all(normalise_text(cell_val), "\\d+\\.?\\d*")[[1]]) else numeric(0)
    out_nums_all <- as.numeric(unlist(lapply(names(parsed), function(nm) {
      v <- parsed[[nm]]; if (is.numeric(v) && !is.na(v)) v else NULL
    })))
    out_nums_all <- out_nums_all[!is.na(out_nums_all)]
    unused_nums  <- all_raw_nums[!sapply(all_raw_nums, function(n) any(abs(n - out_nums_all) < 0.001, na.rm = TRUE))]
    flag_unused  <- if (length(unused_nums) > 0) "Y" else "N"

    text_remainder <- if (!is.na(cell_val)) {
      r <- str_remove_all(normalise_text(cell_val), "\\d+\\.?\\d*")
      r <- str_remove_all(r, "[()\\[\\]%/;:,.\\u00b1+~\\-]")
      r <- str_remove_all(r, regex(
        "\\b(days?|hours?|hrs?|wks?|weeks?|SD|IQR|mean|median|medican|range|average|overall|and|or|the|to|from|with|was|were|of|n|vs\\.?)\\b",
        ignore_case = TRUE))
      str_trim(str_replace_all(r, "\\s+", " "))
    } else ""
    flag_additional_text <- if (nchar(text_remainder) > 3) "Y" else "N"

    # FIX L: flag cells that have text content but no numeric data at all
    flag_text_only <- if (!is.na(cell_val) && nchar(str_trim(cell_val)) > 0 &&
                          !str_detect(cell_val, "\\d")) "Y" else "N"

    flag_dup <- "N"
    if (COL_AGE %in% colnames(df) && !is.na(cell_val)) {
      age_raw <- scalar_val(row[[COL_AGE]], "character")
      if (!is.na(age_raw)) {
        cell_rnd <- round(as.numeric(str_extract_all(normalise_text(cell_val),  "\\d+\\.?\\d*")[[1]]), 4)
        age_rnd  <- round(as.numeric(str_extract_all(normalise_text(age_raw),   "\\d+\\.?\\d*")[[1]]), 4)
        if (length(cell_rnd) > 0 && setequal(cell_rnd, age_rnd)) flag_dup <- "Y"
      }
    }

    flag_unspecified <- if (isTRUE(parsed$overall_avg_not_specified == "Y") ||
                            isTRUE(parsed$overall_unspec_reported   == "Y") ||
                            isTRUE(parsed$group_1_avg_not_specified == "Y") ||
                            isTRUE(parsed$group_1_unspec_reported   == "Y") ||
                            isTRUE(parsed$group_2_avg_not_specified == "Y") ||
                            isTRUE(parsed$group_2_unspec_reported   == "Y")) "Y" else "N"

    g1_present <- !is.na(parsed$group_1_name) || !is.na(parsed$group_1_mean_value) ||
      !is.na(parsed$group_1_median_value)      || !is.na(parsed$group_1_avg_not_specified_value)
    g2_present <- !is.na(parsed$group_2_name) || !is.na(parsed$group_2_mean_value) ||
      !is.na(parsed$group_2_median_value)      || !is.na(parsed$group_2_avg_not_specified_value)
    cell_has_subgroups <- isTRUE(parsed$subgroups_reported == "Y") || g1_present

    if (cell_has_subgroups) {
      final_has_sub <- TRUE
      final_n_sub   <- as.integer(g1_present) + as.integer(g2_present)
      final_n_sub   <- max(final_n_sub, 1L)
    } else {
      final_has_sub <- sub_info$has_subgroups
      final_n_sub   <- sub_info$n_subgroups
    }

    flag_excess_subgroups <- if (!is.na(final_n_sub) && final_n_sub > 2) "Y" else "N"

    tag_name <- function(nm) {
      if (is.na(nm) || str_detect(nm, regex("^group_\\d+$", ignore_case = TRUE))) nm
      else paste0(nm, " [", current_tag, "]")
    }
    cross_names <- get_cell_level_names_for_study(study_num)
    all_subgroup_names <- combine_subgroup_names(
      sub_info$names_from_T, sub_info$name_from_BB, sub_info$name_from_BC,
      tag_name(parsed$group_1_name), tag_name(parsed$group_2_name), cross_names)

    yn_has  <- function(v)    if (!is.na(v)) "Y" else "N"
    yn_miss <- function(...) if (all(is.na(c(...)))) "Y" else "N"

    individual_flags <- c(
      flag_not                        = flag_not,
      flag_question                   = flag_question,
      flag_proportion_data            = flag_prop_data,
      flag_unused_numbers             = flag_unused,
      flag_additional_text            = flag_additional_text,
      flag_text_only                  = flag_text_only,          # FIX L
      flag_duplication                = flag_dup,
      flag_unspecified                = flag_unspecified,
      overall_flag_approx             = parsed$overall_flag_approx,
      overall_flag_UQR_lt_LQR         = parsed$overall_flag_UQR_lt_LQR,
      overall_flag_IQR_with_pm        = parsed$overall_flag_IQR_with_pm,
      overall_flag_unparsed_numbers   = parsed$overall_flag_unparsed_numbers,
      overall_flag_inferred_mean      = parsed$overall_flag_inferred_mean,   # FIX K
      overall_flag_inferred_median    = parsed$overall_flag_inferred_median, # FIX K
      group_1_flag_approx             = parsed$group_1_flag_approx,
      group_1_flag_UQR_lt_LQR         = parsed$group_1_flag_UQR_lt_LQR,
      group_1_flag_IQR_with_pm        = parsed$group_1_flag_IQR_with_pm,
      group_1_flag_unparsed_numbers   = parsed$group_1_flag_unparsed_numbers,
      group_1_flag_inferred_mean      = parsed$group_1_flag_inferred_mean,   # FIX K
      group_1_flag_inferred_median    = parsed$group_1_flag_inferred_median, # FIX K
      group_2_flag_approx             = parsed$group_2_flag_approx,
      group_2_flag_UQR_lt_LQR         = parsed$group_2_flag_UQR_lt_LQR,
      group_2_flag_IQR_with_pm        = parsed$group_2_flag_IQR_with_pm,
      group_2_flag_unparsed_numbers   = parsed$group_2_flag_unparsed_numbers,
      group_2_flag_inferred_mean      = parsed$group_2_flag_inferred_mean,   # FIX K
      group_2_flag_inferred_median    = parsed$group_2_flag_inferred_median  # FIX K
    )
    flag_any        <- if (any(individual_flags == "Y", na.rm = TRUE)) "Y" else "N"
    triggered       <- names(individual_flags)[!is.na(individual_flags) & individual_flags == "Y"]
    any_flag_string <- if (length(triggered) > 0) paste(triggered, collapse = ", ") else NA_character_

    # FIX I: overall_reported and subgroups_reported removed from output;
    #        group_1_name / group_2_name moved adjacent to their data blocks
    rows_out[[i]] <- list(
      # ---- Basic ----
      study_number                          = study_num,
      study_reference                       = study_ref,
      original_observation                  = cell_val,
      # ---- Flags ----
      flag_any                              = flag_any,
      any_flag_string                       = any_flag_string,
      flag_not                              = flag_not,
      flag_question                         = flag_question,
      flag_proportion_data                  = flag_prop_data,
      flag_unused_numbers                   = flag_unused,
      flag_additional_text                  = flag_additional_text,
      flag_text_only                        = flag_text_only,          # FIX L
      flag_duplication                      = flag_dup,
      flag_unspecified                      = flag_unspecified,
      has_subgroups                         = if (isTRUE(final_has_sub)) "Y" else "N",
      n_subgroups                           = final_n_sub,
      flag_excess_subgroups                 = flag_excess_subgroups,
      all_subgroup_names                    = all_subgroup_names,
      # ---- Overall average ----
      overall_mean_reported                 = parsed$overall_mean_reported,
      overall_mean_value                    = parsed$overall_mean_value,
      overall_median_reported               = parsed$overall_median_reported,
      overall_median_value                  = parsed$overall_median_value,
      overall_avg_not_specified             = parsed$overall_avg_not_specified,
      overall_avg_not_specified_value       = parsed$overall_avg_not_specified_value,
      overall_avg_not_reported              = parsed$overall_avg_not_reported,
      # ---- Overall dispersion ----
      overall_SD_reported                   = parsed$overall_SD_reported,
      overall_SD_value                      = parsed$overall_SD_value,
      overall_IQR_reported                  = parsed$overall_IQR_reported,
      overall_IQR_value                     = parsed$overall_IQR_value,
      overall_unspec_disp_reported          = parsed$overall_unspec_reported,
      overall_unspec_disp_value             = parsed$overall_unspec_disp_value,
      overall_dispersion_not_reported       = parsed$overall_dispersion_not_reported,
      # ---- Overall boundary types reported ----
      overall_LQR_reported                  = yn_has(parsed$overall_IQR_LQR),
      overall_UQR_reported                  = yn_has(parsed$overall_IQR_UQR),
      overall_lower_range_reported          = yn_has(parsed$overall_range_lower),
      overall_upper_range_reported          = yn_has(parsed$overall_range_upper),
      overall_unspec_lower_reported         = yn_has(parsed$overall_unspec_lower),
      overall_unspec_upper_reported         = yn_has(parsed$overall_unspec_upper),
      overall_lower_boundary_not_reported   = yn_miss(parsed$overall_IQR_LQR,   parsed$overall_range_lower,  parsed$overall_unspec_lower),
      overall_upper_boundary_not_reported   = yn_miss(parsed$overall_IQR_UQR,   parsed$overall_range_upper,  parsed$overall_unspec_upper),
      # ---- Overall boundary values ----
      overall_IQR_LQR                       = parsed$overall_IQR_LQR,
      overall_IQR_UQR                       = parsed$overall_IQR_UQR,
      overall_range_lower                   = parsed$overall_range_lower,
      overall_range_upper                   = parsed$overall_range_upper,
      overall_unspec_lower                  = parsed$overall_unspec_lower,
      overall_unspec_upper                  = parsed$overall_unspec_upper,
      # ---- Overall other flags ----
      overall_flag_approx                   = parsed$overall_flag_approx,
      overall_flag_UQR_lt_LQR               = parsed$overall_flag_UQR_lt_LQR,
      overall_flag_IQR_with_pm              = parsed$overall_flag_IQR_with_pm,
      overall_flag_unparsed_numbers         = parsed$overall_flag_unparsed_numbers,
      overall_flag_inferred_mean            = parsed$overall_flag_inferred_mean,   # FIX K
      overall_flag_inferred_median          = parsed$overall_flag_inferred_median, # FIX K
      # ---- Group 1 (name adjacent to data) ----
      group_1_name                          = parsed$group_1_name,
      group_1_mean_reported                 = parsed$group_1_mean_reported,
      group_1_mean_value                    = parsed$group_1_mean_value,
      group_1_median_reported               = parsed$group_1_median_reported,
      group_1_median_value                  = parsed$group_1_median_value,
      group_1_avg_not_specified             = parsed$group_1_avg_not_specified,
      group_1_avg_not_specified_value       = parsed$group_1_avg_not_specified_value,
      group_1_avg_not_reported              = parsed$group_1_avg_not_reported,
      # ---- Group 1 dispersion ----
      group_1_SD_reported                   = parsed$group_1_SD_reported,
      group_1_SD_value                      = parsed$group_1_SD_value,
      group_1_IQR_reported                  = parsed$group_1_IQR_reported,
      group_1_IQR_value                     = parsed$group_1_IQR_value,
      group_1_unspec_disp_reported          = parsed$group_1_unspec_reported,
      group_1_unspec_disp_value             = parsed$group_1_unspec_disp_value,
      group_1_dispersion_not_reported       = parsed$group_1_dispersion_not_reported,
      # ---- Group 1 boundary types reported ----
      group_1_LQR_reported                  = yn_has(parsed$group_1_IQR_LQR),
      group_1_UQR_reported                  = yn_has(parsed$group_1_IQR_UQR),
      group_1_lower_range_reported          = yn_has(parsed$group_1_range_lower),
      group_1_upper_range_reported          = yn_has(parsed$group_1_range_upper),
      group_1_unspec_lower_reported         = yn_has(parsed$group_1_unspec_lower),
      group_1_unspec_upper_reported         = yn_has(parsed$group_1_unspec_upper),
      group_1_lower_boundary_not_reported   = yn_miss(parsed$group_1_IQR_LQR,   parsed$group_1_range_lower,  parsed$group_1_unspec_lower),
      group_1_upper_boundary_not_reported   = yn_miss(parsed$group_1_IQR_UQR,   parsed$group_1_range_upper,  parsed$group_1_unspec_upper),
      # ---- Group 1 boundary values ----
      group_1_IQR_LQR                       = parsed$group_1_IQR_LQR,
      group_1_IQR_UQR                       = parsed$group_1_IQR_UQR,
      group_1_range_lower                   = parsed$group_1_range_lower,
      group_1_range_upper                   = parsed$group_1_range_upper,
      group_1_unspec_lower                  = parsed$group_1_unspec_lower,
      group_1_unspec_upper                  = parsed$group_1_unspec_upper,
      # ---- Group 1 other flags ----
      group_1_flag_approx                   = parsed$group_1_flag_approx,
      group_1_flag_UQR_lt_LQR               = parsed$group_1_flag_UQR_lt_LQR,
      group_1_flag_IQR_with_pm              = parsed$group_1_flag_IQR_with_pm,
      group_1_flag_unparsed_numbers         = parsed$group_1_flag_unparsed_numbers,
      group_1_flag_inferred_mean            = parsed$group_1_flag_inferred_mean,   # FIX K
      group_1_flag_inferred_median          = parsed$group_1_flag_inferred_median, # FIX K
      # ---- Group 2 (name adjacent to data) ----
      group_2_name                          = parsed$group_2_name,
      group_2_mean_reported                 = parsed$group_2_mean_reported,
      group_2_mean_value                    = parsed$group_2_mean_value,
      group_2_median_reported               = parsed$group_2_median_reported,
      group_2_median_value                  = parsed$group_2_median_value,
      group_2_avg_not_specified             = parsed$group_2_avg_not_specified,
      group_2_avg_not_specified_value       = parsed$group_2_avg_not_specified_value,
      group_2_avg_not_reported              = parsed$group_2_avg_not_reported,
      # ---- Group 2 dispersion ----
      group_2_SD_reported                   = parsed$group_2_SD_reported,
      group_2_SD_value                      = parsed$group_2_SD_value,
      group_2_IQR_reported                  = parsed$group_2_IQR_reported,
      group_2_IQR_value                     = parsed$group_2_IQR_value,
      group_2_unspec_disp_reported          = parsed$group_2_unspec_reported,
      group_2_unspec_disp_value             = parsed$group_2_unspec_disp_value,
      group_2_dispersion_not_reported       = parsed$group_2_dispersion_not_reported,
      # ---- Group 2 boundary types reported ----
      group_2_LQR_reported                  = yn_has(parsed$group_2_IQR_LQR),
      group_2_UQR_reported                  = yn_has(parsed$group_2_IQR_UQR),
      group_2_lower_range_reported          = yn_has(parsed$group_2_range_lower),
      group_2_upper_range_reported          = yn_has(parsed$group_2_range_upper),
      group_2_unspec_lower_reported         = yn_has(parsed$group_2_unspec_lower),
      group_2_unspec_upper_reported         = yn_has(parsed$group_2_unspec_upper),
      group_2_lower_boundary_not_reported   = yn_miss(parsed$group_2_IQR_LQR,   parsed$group_2_range_lower,  parsed$group_2_unspec_lower),
      group_2_upper_boundary_not_reported   = yn_miss(parsed$group_2_IQR_UQR,   parsed$group_2_range_upper,  parsed$group_2_unspec_upper),
      # ---- Group 2 boundary values ----
      group_2_IQR_LQR                       = parsed$group_2_IQR_LQR,
      group_2_IQR_UQR                       = parsed$group_2_IQR_UQR,
      group_2_range_lower                   = parsed$group_2_range_lower,
      group_2_range_upper                   = parsed$group_2_range_upper,
      group_2_unspec_lower                  = parsed$group_2_unspec_lower,
      group_2_unspec_upper                  = parsed$group_2_unspec_upper,
      # ---- Group 2 other flags ----
      group_2_flag_approx                   = parsed$group_2_flag_approx,
      group_2_flag_UQR_lt_LQR               = parsed$group_2_flag_UQR_lt_LQR,
      group_2_flag_IQR_with_pm              = parsed$group_2_flag_IQR_with_pm,
      group_2_flag_unparsed_numbers         = parsed$group_2_flag_unparsed_numbers,
      group_2_flag_inferred_mean            = parsed$group_2_flag_inferred_mean,   # FIX K
      group_2_flag_inferred_median          = parsed$group_2_flag_inferred_median  # FIX K
    )
  }

  all_sheet_data[[sheet_name]] <- bind_rows(rows_out)
}

# =============================================================================
# PART 4: Write full and NR workbooks
# =============================================================================
dir.create(dirname(output_path), showWarnings = FALSE, recursive = TRUE)

FLAG_COLS <- c(
  "flag_any","any_flag_string","flag_not","flag_question","flag_proportion_data",
  "flag_unused_numbers","flag_additional_text","flag_text_only",   # FIX L: added flag_text_only
  "flag_duplication","flag_unspecified",
  "has_subgroups","n_subgroups","flag_excess_subgroups","all_subgroup_names"
)

write_workbook <- function(sheet_data_list, path, fill_nr = FALSE) {
  wb           <- createWorkbook()
  header_style <- createStyle(fontName = "Arial", fontSize = 10, textDecoration = "bold",
                              fgFill = "#D9E1F2", wrapText = TRUE, border = "Bottom")
  flag_style   <- createStyle(fgFill = "#FFF2CC")

  for (sname in names(sheet_data_list)) {
    df_sheet <- sheet_data_list[[sname]]
    if (fill_nr) {
      df_sheet <- df_sheet %>%
        mutate(across(everything(), ~ { x <- .x; x[is.na(x) | as.character(x) == ""] <- "NR"; x }))
    }
    addWorksheet(wb, sheetName = substr(sname, 1, 31))
    writeData(wb, sheet = sname, x = df_sheet, headerStyle = header_style)
    n_rows <- nrow(df_sheet)
    if (n_rows > 0) {
      for (j in seq_along(colnames(df_sheet))) {
        if (colnames(df_sheet)[j] %in% FLAG_COLS)
          addStyle(wb, sheet = sname, style = flag_style,
                   rows = 2:(n_rows + 1), cols = j, gridExpand = TRUE)
      }
    }
    for (j in seq_along(colnames(df_sheet)))
      setColWidths(wb, sheet = sname, cols = j,
                   widths = max(12, nchar(colnames(df_sheet)[j]) + 2))
  }
  saveWorkbook(wb, path, overwrite = TRUE)
  cat("Saved:", path, "\n")
}

write_workbook(all_sheet_data, output_path,    fill_nr = FALSE)
write_workbook(all_sheet_data, output_path_nr, fill_nr = TRUE)

# =============================================================================
# PART 5: Write condensed workbook
# =============================================================================
CONDENSED_COLS <- c(
  "study_number", "study_reference", "original_observation",
  "any_flag_string", "has_subgroups", "n_subgroups", "all_subgroup_names",
  "overall_mean_value", "overall_median_value", "overall_avg_not_specified_value",
  "overall_SD_value", "overall_IQR_value", "overall_unspec_disp_value",
  "overall_IQR_LQR", "overall_IQR_UQR",
  "overall_range_lower", "overall_range_upper",
  "overall_unspec_lower", "overall_unspec_upper",
  "group_1_name",
  "group_1_mean_value", "group_1_median_value", "group_1_avg_not_specified_value",
  "group_1_SD_value", "group_1_IQR_value", "group_1_unspec_disp_value",
  "group_1_IQR_LQR", "group_1_IQR_UQR",
  "group_1_range_lower", "group_1_range_upper",
  "group_1_unspec_lower", "group_1_unspec_upper",
  "group_2_name",
  "group_2_mean_value", "group_2_median_value", "group_2_avg_not_specified_value",
  "group_2_SD_value", "group_2_IQR_value", "group_2_unspec_disp_value",
  "group_2_IQR_LQR", "group_2_IQR_UQR",
  "group_2_range_lower", "group_2_range_upper",
  "group_2_unspec_lower", "group_2_unspec_upper"
)

condensed_data <- lapply(all_sheet_data, function(df_sheet) {
  present <- intersect(CONDENSED_COLS, colnames(df_sheet))
  missing <- setdiff(CONDENSED_COLS, colnames(df_sheet))
  if (length(missing) > 0)
    warning("Condensed workbook missing expected column(s): ", paste(missing, collapse = ", "))
  df_sheet[, present, drop = FALSE]
})
write_workbook(condensed_data, output_path_condensed, fill_nr = FALSE)

cat("Done.\n")
