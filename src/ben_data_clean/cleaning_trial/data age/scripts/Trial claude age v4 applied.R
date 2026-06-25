# Jesus is LORD!

# From claude (see separate doc for methodology of using claude)

# Addition at line 49 (to load readxl package, and then open file)
# Apply in part 4 and 5 (changing reference terms and file names, etc)


# =============================================================================
# Age Data Parsing Script — Version 4
# =============================================================================
# Major changes vs v3:
#
# COLUMN RESTRUCTURE (per user request):
#   Average now split into: mean_reported (Y/N), mean_value,
#                           median_reported (Y/N), median_value,
#                           avg_not_specified (Y/N), avg_not_specified_value,
#                           avg_not_reported (Y/N)
#   Dispersion now split into: SD_reported (Y/N), SD_value,
#                              IQR_reported (Y/N), IQR_LQR, IQR_UQR,
#                              range_reported (Y/N), range_lower, range_upper,
#                              unspec_reported (Y/N), unspec_lower, unspec_upper,
#                              dispersion_not_reported (Y/N)
#   Two output sheets: one with NA as blank, one with NA as "not reported"
#
# LOGIC FIXES:
#   - "average" keyword no longer maps to mean; maps to avg_not_specified
#   - "not_reported" properly distinguished from "unknown/not specified"
#   - SD regex no longer captures the mean value (root cause of #2, #8, #18, #79, #93)
#   - "(SD X)" now correctly parsed (#7, #34)
#   - Category prefix stripping no longer strips median/IQR values (#14, #16, #17)
#   - Row 6 no longer split as 2 groups (value semicolons correctly detected)
#   - Row 72 subgroups now pulled (indigenous, non-indigenous)
#   - Row 78 category-only cells no longer produce false average
#   - Row 92 no longer split as 2 groups
#   - Group avg type defaults to "not_reported" when not stated (not blank)
#   - Category shared to both group_1 and group_2 when same for both
#   - Group names extracted more reliably without hardcoded values
#   - Row 51 category now pulled
#   - Row 104 median secondary stat now also pulled
#   - Rows 50, 56, 69, 75, 82, 90, 107, 108, 112 various fixes
#
# REPLACE "my_tibble" and "age_column" with your actual names.
# =============================================================================

library(dplyr)
library(stringr)
library(openxlsx)
library(readxl)
age <- read_xlsx("C:/Users/bgodd/Documents/age.xlsx")

# =============================================================================
# HELPER: normalise_text()
# =============================================================================
# Standardises special unicode characters to ASCII equivalents

normalise_text <- function(text) {
  if (is.na(text)) return(NA_character_)
  text <- str_replace_all(text, "\u2013|\u2014", "-")   # en/em dash -> hyphen
  text <- str_replace_all(text, "\u00b1", "±")           # ± normalise
  text <- str_replace_all(text, "\u2265", ">=")          # ≥ -> >=
  text <- str_replace_all(text, "\u2264", "<=")          # ≤ -> <=
  text <- str_replace_all(text, "\u223c", "~")           # ∼ -> ~
  str_trim(text)
}


# =============================================================================
# HELPER: extract_category()
# =============================================================================
# Extracts age category label from text.
# Handles: "adult", ">=18 years", "18-65y", "grouped ages...", "Adults>=18y"
# FIX (#51): also captures "grouped as <40, 40-64..." style descriptions
# FIX (#78): "Adults>=18y" is category only — the "18" is NOT an average age

extract_category <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(NA_character_)
  
  m <- str_extract(text, regex(paste0(
    # "age (grouped as ...)" full phrase
    "age\\s*\\([^)]*grouped[^)]*\\)|",
    # "grouped ages..." standalone
    "grouped ages?[^;.]*|",
    # "Adults>=18y" or "adult >= 18 years"
    "adults?\\s*[>=<]+\\s*\\d+\\s*y(ears?)?|",
    # standalone "adult/adults/elderly/paediatric/pediatric/child/adolescent"
    "\\badults?\\b|\\belderly\\b|\\bpaediatric\\b|\\bpediatric\\b|",
    "\\bchild\\b|\\badolescent\\b|\\bgeriatric\\b|",
    # ">= 18 years" / ">80y" style thresholds
    "[>=<]+\\s*\\d+\\s*y(ears?)?|",
    # "18-65y" style ranges
    "\\d+\\s*-\\s*\\d+\\s*y(ears?)?"
  ), ignore_case = TRUE))
  
  if (is.na(m)) return(NA_character_)
  str_trim(m)
}


# =============================================================================
# HELPER: is_category_only()
# =============================================================================
# Returns TRUE if the cell contains ONLY a category label (no numeric age data)
# Used to set avg_not_reported = Y and prevent false average extraction
# FIX (#78, #9, #11, #13, #48, #64, etc.)

is_category_only <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(TRUE)
  t <- str_trim(text)
  # Remove the category part and see if any numeric age data remains
  stripped <- str_remove(t, regex(paste0(
    "age\\s*\\([^)]*grouped[^)]*\\)|",
    "adults?\\s*[>=<]+\\s*\\d+\\s*y(ears?)?\\s*:?|",
    "\\badults?\\b\\s*:?|\\belderly\\b\\s*:?|",
    "[>=<]+\\s*\\d+\\s*y(ears?)?\\s*:?|",
    "\\d+\\s*-\\s*\\d+\\s*y(ears?)?\\s*:?"
  ), ignore_case = TRUE))
  stripped <- str_trim(stripped)
  # If nothing meaningful remains after removing the category label
  nchar(stripped) == 0 || !str_detect(stripped, "\\d")
}


# =============================================================================
# HELPER: is_value_semicolon()
# =============================================================================
# Returns TRUE if semicolons in the cell separate values (not groups)
# FIX (#6): "mean years 65.4; SD 14.6; median years 67" must NOT be split
# FIX (#14): "(Q1 ; Q3)" must NOT trigger group split

is_value_semicolon <- function(text) {
  if (!str_detect(text, ";")) return(FALSE)
  # Q1;Q3 pattern is always a value separator
  if (str_detect(text, regex("Q1\\s*;\\s*Q3", ignore_case = TRUE))) return(TRUE)
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  # Each segment is value-like if it is a stat keyword + optional number,
  # but NOT a group-name + number (groups have longer descriptive names)
  value_like <- sapply(segs, function(s) {
    is_stat_fragment <- str_detect(s, regex(
      "^\\s*(SD|IQR|Q[123]|mean|median|average|range|years?|y|\\d+\\.?\\d*)\\b",
      ignore_case = TRUE
    ))
    has_group_name <- str_detect(s, regex(
      "^\\s*[a-zA-Z]{3,}\\s+[a-zA-Z]{2,}.*\\d",  # two+ words before a number
      ignore_case = FALSE
    ))
    is_stat_fragment && !has_group_name
  })
  all(value_like)
}


# =============================================================================
# HELPER: strip_cell_header()
# =============================================================================
# Strips a leading "format header" from the cell BEFORE group splitting.
# e.g. ">=18 years: years median IQR; group1 X; group2 Y"
#       -> strips ">=18 years: years median IQR; "
#       -> leaves "group1 X; group2 Y"
#
# FIX (#10, #15, #20): category prefix + format words no longer become group_1
# CRITICAL: only strips UP TO the first semicolon if that first segment looks
# like a header (no group-name-like content), not actual data.
#
# Returns list(header=..., remainder=...) so header can be parsed separately

strip_cell_header <- function(text) {
  result <- list(header = NA_character_, remainder = text)
  if (!str_detect(text, ";")) return(result)
  
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  if (length(segs) < 2) return(result)
  
  first <- segs[1]
  
  # The first segment is a header if:
  # 1. It contains a category marker (>=, >, adult, etc.)
  # 2. AND it contains format words (mean, median, IQR, SD, years) but NO standalone number
  #    that could be an actual age value
  has_category    <- str_detect(first, regex("[>=<]+\\s*\\d+|\\badults?\\b", ignore_case = TRUE))
  has_format_word <- str_detect(first, regex("\\b(mean|median|IQR|SD|years?|y)\\b", ignore_case = TRUE))
  has_age_number  <- str_detect(first, regex("\\b(\\d{2,3}(\\.\\d+)?)\\s*[y\\(\\[±]", ignore_case = TRUE))
  
  if (has_category && has_format_word && !has_age_number) {
    result$header    <- first
    result$remainder <- paste(segs[-1], collapse = ";")
  }
  result
}


# =============================================================================
# PART 1: parse_age_block()
# =============================================================================
# Restructured to produce the new column layout.
# Returns a named list with all average and dispersion fields as Y/N flags + values.

parse_age_block <- function(text,
                            inherit_mean   = NULL,
                            inherit_median = NULL,
                            inherit_disp   = NULL) {
  
  # Initialise all outputs
  out <- list(
    category               = NA_character_,
    mean_reported          = NA_character_,
    mean_value             = NA_real_,
    median_reported        = NA_character_,
    median_value           = NA_real_,
    avg_not_specified      = NA_character_,
    avg_not_specified_value= NA_real_,
    avg_not_reported       = NA_character_,
    SD_reported            = NA_character_,
    SD_value               = NA_real_,
    IQR_reported           = NA_character_,
    IQR_LQR                = NA_real_,
    IQR_UQR                = NA_real_,
    range_reported         = NA_character_,
    range_lower            = NA_real_,
    range_upper            = NA_real_,
    unspec_reported        = NA_character_,
    unspec_lower           = NA_real_,
    unspec_upper           = NA_real_,
    dispersion_not_reported= NA_character_
  )
  
  if (is.na(text) || str_trim(text) == "") {
    out$avg_not_reported        <- "Y"
    out$dispersion_not_reported <- "Y"
    return(out)
  }
  
  text <- normalise_text(text)
  
  # --- CATEGORY ---
  out$category <- extract_category(text)
  
  # --- SHORT CIRCUIT: category-only cells ---
  # FIX (#78, #9, #11, #13, #48, etc.): no avg data in these cells
  if (is_category_only(text)) {
    out$avg_not_reported        <- "Y"
    out$mean_reported           <- "N"
    out$median_reported         <- "N"
    out$avg_not_specified       <- "N"
    out$dispersion_not_reported <- "Y"
    out$SD_reported             <- "N"
    out$IQR_reported            <- "N"
    out$range_reported          <- "N"
    out$unspec_reported         <- "N"
    return(out)
  }
  
  # -------------------------------------------------------------------------
  # AVERAGE EXTRACTION
  # Strategy: find mean value and median value INDEPENDENTLY.
  # FIX (#88): "average" no longer maps to mean; maps to avg_not_specified
  # FIX (#104): both mean AND median extracted when both present
  # FIX (#2, #8, #18): mean regex anchored so it does NOT grab SD numbers
  # -------------------------------------------------------------------------
  
  # Remove secondary "(median, X)" notation before checking for primary avg type
  # FIX (#104): we extract it first as median_secondary, then remove for mean search
  median_secondary <- NA_real_
  secondary_match <- str_extract(text, regex("\\(\\s*median\\s*,?\\s*(\\d+\\.?\\d*)\\s*\\)", ignore_case = TRUE))
  if (!is.na(secondary_match)) {
    median_secondary <- as.numeric(str_extract(secondary_match, "\\d+\\.?\\d*$"))
  }
  text_no_secondary <- str_remove(text, regex("\\(\\s*median\\s*,?\\s*\\d+\\.?\\d*\\s*\\)", ignore_case = TRUE))
  
  # --- MEAN ---
  mean_val <- NA_real_
  has_mean_keyword <- str_detect(text_no_secondary, regex("\\bmean\\b", ignore_case = TRUE))
  
  if (has_mean_keyword) {
    # Pattern: mean/mean age/mean age was/mean age at admission of... NUMBER
    m1 <- str_extract(text_no_secondary, regex(
      "\\bmean\\s*(?:age(?:d|s)?)?\\s*(?:(?:was|of|at admission of|,)\\s*)?[~=:]?\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m1)) mean_val <- as.numeric(str_extract(m1, "\\d+\\.?\\d*$"))
    
    # Pattern: "Mean 75", "Mean 49.5y" — keyword then immediate number
    if (is.na(mean_val)) {
      m2 <- str_extract(text_no_secondary, regex("\\bmean\\s+(\\d+\\.?\\d*)\\s*y?\\b", ignore_case = TRUE))
      if (!is.na(m2)) mean_val <- as.numeric(str_extract(m2, "\\d+\\.?\\d*"))
    }
    
    # Pattern: "63 (mean)" — number before the keyword in brackets
    if (is.na(mean_val)) {
      m3 <- str_extract(text_no_secondary, regex("(\\d+\\.?\\d*)\\s*\\(\\s*mean\\s*\\)", ignore_case = TRUE))
      if (!is.na(m3)) mean_val <- as.numeric(str_extract(m3, "^\\d+\\.?\\d*"))
    }
    
    # Pattern: "Mean (SD): 55.5" or "mean (SD) age was 67.5" — number AFTER label block
    # FIX (#79, #84, #93): "Mean (SD) X" — X is the mean, bracket content is label
    if (is.na(mean_val)) {
      m4 <- str_extract(text_no_secondary, regex(
        "\\bmean\\s*\\([^)]*\\)\\s*(?:age\\s*was\\s*)?[=:]?\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m4)) mean_val <- as.numeric(str_extract(m4, "\\d+\\.?\\d*$"))
    }
    
    # FIX (#3): fallback — first standalone number after stripping keyword and label
    if (is.na(mean_val)) {
      stripped_for_num <- str_remove(text_no_secondary, regex(
        ".*?\\bmean\\b\\s*(?:age(?:d|s)?)?\\s*(?:(?:was|of|at admission|were|aged)\\s*)?[=:,]?\\s*",
        ignore_case = TRUE))
      # Remove any bracketed label like "(SD)" or "(± SD)" before the number
      stripped_for_num <- str_remove(stripped_for_num, regex("^\\s*\\([^\\d)]*\\)\\s*", ignore_case = TRUE))
      m5 <- str_extract(stripped_for_num, "^\\s*(\\d+\\.?\\d*)")
      if (!is.na(m5)) mean_val <- as.numeric(str_extract(m5, "\\d+\\.?\\d*"))
    }
  }
  
  # --- MEDIAN ---
  median_val <- NA_real_
  has_median_keyword <- str_detect(text, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))
  
  if (has_median_keyword) {
    m1 <- str_extract(text, regex(
      "\\b(?:median|medican)\\s*(?:age(?:\\s*(?:of|,|was|at))?)?\\s*[~=:,]?\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m1)) {
      val <- as.numeric(str_extract(m1, "\\d+\\.?\\d*$"))
      # FIX (#14, #16, #17): validate the number is plausible as an age (5-120)
      if (!is.na(val) && val >= 5 && val <= 120) median_val <- val
    }
    
    # Pattern: "Median ~65" — approximate value
    if (is.na(median_val)) {
      m2 <- str_extract(text, regex("\\bmedian\\b\\s*~\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m2)) median_val <- as.numeric(str_extract(m2, "\\d+\\.?\\d*$"))
    }
  }
  
  # If secondary median stat present and no primary median found yet, use it
  if (!is.na(median_secondary) && is.na(median_val)) {
    median_val <- median_secondary
    has_median_keyword <- TRUE
  } else if (!is.na(median_secondary)) {
    # Both primary and secondary median present — keep both (primary already set)
    # secondary recorded as median_value below
  }
  
  # --- AVERAGE NOT SPECIFIED (bare number or "average" keyword) ---
  # FIX (#88): "average" -> avg_not_specified, not mean
  avg_ns_val <- NA_real_
  has_average_keyword <- str_detect(text, regex("\\baverage\\b", ignore_case = TRUE))
  
  if (!has_mean_keyword && !has_median_keyword) {
    # Fallback: bare number before brackets or ± or end of string
    m_bare <- str_extract(text, "^[^\\d]*(\\d+\\.?\\d*)\\s*y?\\s*(?:[\\(\\[±\\+]|$)")
    if (!is.na(m_bare)) {
      val <- as.numeric(str_extract(m_bare, "\\d+\\.?\\d*"))
      # Validate: must be plausible age (5-120) and not just a category number
      cat_num <- str_extract(out$category %||% "", "\\d+")
      if (!is.na(val) && val >= 5 && val <= 120) {
        if (is.na(cat_num) || as.numeric(cat_num) != val) {
          avg_ns_val <- val
        }
      }
    }
  } else if (has_average_keyword && !has_mean_keyword) {
    # "average age was X" -> avg_not_specified
    m_avg <- str_extract(text, regex(
      "\\baverage\\s*(?:age\\s*(?:was|of)?)?\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m_avg)) avg_ns_val <- as.numeric(str_extract(m_avg, "\\d+\\.?\\d*$"))
  }
  
  # --- ASSIGN AVERAGE FLAGS ---
  if (!is.na(mean_val) || has_mean_keyword) {
    out$mean_reported <- "Y"
    out$mean_value    <- mean_val
  } else {
    out$mean_reported <- "N"
  }
  
  if (!is.na(median_val) || has_median_keyword) {
    out$median_reported <- "Y"
    out$median_value    <- if (!is.na(median_val)) median_val else
      if (!is.na(median_secondary)) median_secondary else NA_real_
  } else {
    out$median_reported <- "N"
  }
  
  if (!is.na(avg_ns_val) || has_average_keyword) {
    out$avg_not_specified       <- "Y"
    out$avg_not_specified_value <- avg_ns_val
  } else {
    out$avg_not_specified <- "N"
  }
  
  # avg_not_reported: Y only if none of the above found
  out$avg_not_reported <- if (out$mean_reported == "N" &&
                              out$median_reported == "N" &&
                              out$avg_not_specified == "N") "Y" else "N"
  
  # Inherit average types from overall if this is a subgroup with no keywords
  # FIX (#50, #56, #72, #75, #107, #108, #112)
  if (out$avg_not_reported == "Y" && !is.null(inherit_mean) && inherit_mean == "Y") {
    out$mean_reported    <- "Y"
    out$avg_not_reported <- "N"
  }
  if (out$avg_not_reported == "Y" && !is.null(inherit_median) && inherit_median == "Y") {
    out$median_reported  <- "Y"
    out$avg_not_reported <- "N"
  }
  
  # -------------------------------------------------------------------------
  # DISPERSION EXTRACTION
  # Completely rewritten to avoid SD regex grabbing the mean value
  # FIX (#2, #8, #18, #79, #93): only extract SD from AFTER the mean number
  # FIX (#7, #34): "(SD 6.5)" and "(SD 16.4)" now correctly extracted
  # -------------------------------------------------------------------------
  
  # Build a "post-average" text — everything after the average number —
  # for SD/dispersion searching. This prevents mean value being captured as SD.
  post_avg_text <- text
  if (!is.na(mean_val)) {
    # Find position of mean number and take everything after it
    mean_pos <- str_locate(text, as.character(mean_val))[1, "end"]
    if (!is.na(mean_pos)) post_avg_text <- substr(text, mean_pos + 1, nchar(text))
  } else if (!is.na(median_val)) {
    med_pos <- str_locate(text, as.character(median_val))[1, "end"]
    if (!is.na(med_pos)) post_avg_text <- substr(text, med_pos + 1, nchar(text))
  } else if (!is.na(avg_ns_val)) {
    ns_pos <- str_locate(text, as.character(avg_ns_val))[1, "end"]
    if (!is.na(ns_pos)) post_avg_text <- substr(text, ns_pos + 1, nchar(text))
  }
  
  has_iqr  <- str_detect(text, regex("\\bIQR\\b|Q1.*Q3|25/?75%?\\s*IQR", ignore_case = TRUE))
  has_range <- str_detect(text, regex("\\brange[d]?\\b", ignore_case = TRUE))
  has_sd_kw <- str_detect(text, regex("\\bSD\\b|±|\\+\\s*/?\\s*-|\\+-", ignore_case = FALSE)) |
    str_detect(text, regex("\\bSD\\b", ignore_case = TRUE))
  
  # --- SD ---
  sd_val <- NA_real_
  
  # "(SD X)" or "(SD) X" — label before number
  # FIX (#7, #34): search full text for this pattern
  m_sd_brk <- str_extract(text, regex("\\(\\s*SD\\s*\\)\\s*[=:]?\\s*(\\d+\\.?\\d*)|\\(\\s*SD\\s+(\\d+\\.?\\d*)\\s*\\)", ignore_case = TRUE))
  if (!is.na(m_sd_brk)) {
    nums <- str_extract_all(m_sd_brk, "\\d+\\.?\\d*")[[1]]
    if (length(nums) > 0) sd_val <- as.numeric(nums[length(nums)])
  }
  
  # "SD X" (unbracketed) — search post_avg_text to avoid grabbing mean
  if (is.na(sd_val) && !has_iqr) {
    m_sd1 <- str_extract(post_avg_text, regex("\\bSD\\b\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m_sd1)) sd_val <- as.numeric(str_extract(m_sd1, "\\d+\\.?\\d*$"))
  }
  
  # "± X" or "+/- X" — search post_avg_text
  if (is.na(sd_val) && !has_iqr) {
    m_sd2 <- str_extract(post_avg_text, regex("(?:±|\\+\\s*/?\\s*-|\\+-|plus\\s*/?\\s*minus)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd2)) sd_val <- as.numeric(str_extract(m_sd2, "\\d+\\.?\\d*$"))
  }
  
  # Single bracketed number with +/- in full text (e.g. "39 (+/- 11)")
  if (is.na(sd_val) && !has_iqr) {
    m_sd3 <- str_extract(text, regex("(?:±|\\+\\s*/?\\s*-|\\+-)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd3)) sd_val <- as.numeric(str_extract(m_sd3, "\\d+\\.?\\d*$"))
  }
  
  # FIX (#22): single bracketed number, no keywords -> SD
  # Only apply when no other dispersion type found
  single_bracket_sd <- NA_real_
  if (!has_iqr && !has_range && !has_sd_kw) {
    m_sb <- str_extract(post_avg_text, "^\\s*\\(?\\s*(\\d+\\.?\\d*)\\s*\\)?\\s*$")
    if (is.na(m_sb)) {
      m_sb2 <- str_extract(post_avg_text, "^\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
      if (!is.na(m_sb2)) {
        val <- as.numeric(str_extract(m_sb2, "\\d+\\.?\\d*"))
        if (!is.na(val) && !str_detect(m_sb2, "[,;\\-]")) single_bracket_sd <- val
      }
    }
  }
  
  # --- IQR ---
  iqr_lo <- NA_real_; iqr_hi <- NA_real_; iqr_single <- NA_real_
  
  if (has_iqr) {
    m_iqr <- str_extract(text, regex(
      "(?:IQR|Q1\\s*[;,]?\\s*Q3|25/?75%?\\s*IQR)\\s*[:\\(\\[]?\\s*(\\d+\\.?\\d*)\\s*(?:[-,;]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m_iqr)) {
      nums <- as.numeric(str_extract_all(m_iqr, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
    }
    if (is.na(iqr_lo)) {
      # Bracketed pair anywhere in text when IQR keyword present
      m_iqr2 <- str_extract(post_avg_text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[,;\\-]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_iqr2)) {
        nums <- as.numeric(str_extract_all(m_iqr2, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
      }
    }
    # FIX (#102): IQR keyword but only single value (e.g. "IQR: 20" or via ±)
    if (is.na(iqr_lo)) {
      m_iqr3 <- str_extract(post_avg_text, regex("(?:±|\\+\\s*/?\\s*-|\\+-|\\bSD\\b|IQR\\s*:\\s*)(\\d+\\.?\\d*)"))
      if (!is.na(m_iqr3)) iqr_single <- as.numeric(str_extract(m_iqr3, "\\d+\\.?\\d*$"))
    }
    # FIX (#75): IQR keyword present but values in "name X (Y)" format
    # Y in brackets is IQR value — already handled above
  }
  
  # --- RANGE ---
  rng_lo <- NA_real_; rng_hi <- NA_real_
  
  if (has_range) {
    m_rng <- str_extract(text, regex(
      "range[d]?\\s*(?:from)?\\s*(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (is.na(m_rng)) {
      # "(range) X (lo-hi)" or "median (range) age X (lo-hi)"
      m_rng2 <- str_extract(text, regex(
        "\\(\\s*range\\s*\\)[^\\d]*(\\d+\\.?\\d*)[^\\d]+(\\d+\\.?\\d*)\\s*(?:to|-|,)\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m_rng2)) {
        nums <- as.numeric(str_extract_all(m_rng2, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 3) { rng_lo <- nums[2]; rng_hi <- nums[3] }
      }
    } else {
      nums <- as.numeric(str_extract_all(m_rng, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
    }
  }
  
  # --- UNSPECIFIED BRACKETED RANGE ---
  unspec_lo <- NA_real_; unspec_hi <- NA_real_
  
  # "58-62" bare hyphenated pair
  m_bare <- str_extract(str_trim(text), "^\\s*(\\d+\\.?\\d*)\\s*-\\s*(\\d+\\.?\\d*)\\s*$")
  if (!is.na(m_bare)) {
    nums <- as.numeric(str_extract_all(m_bare, "\\d+\\.?\\d*")[[1]])
    if (length(nums) == 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
  }
  
  # Bracketed pair with no dispersion keyword
  if (is.na(unspec_lo) && !has_sd_kw && !has_iqr && !has_range) {
    m_unspec <- str_extract(post_avg_text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
    if (!is.na(m_unspec)) {
      nums <- as.numeric(str_extract_all(m_unspec, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
    }
  }
  
  # --- ASSIGN DISPERSION FLAGS ---
  any_disp <- FALSE
  
  if (!is.na(sd_val) || (!is.na(single_bracket_sd) && !has_iqr && !has_range)) {
    out$SD_reported <- "Y"
    out$SD_value    <- if (!is.na(sd_val)) sd_val else single_bracket_sd
    any_disp <- TRUE
  } else { out$SD_reported <- "N" }
  
  if (has_iqr) {
    out$IQR_reported <- "Y"
    out$IQR_LQR      <- iqr_lo
    out$IQR_UQR      <- iqr_hi
    if (!is.na(iqr_single) && is.na(iqr_lo)) out$SD_value <- iqr_single  # store single IQR value
    any_disp <- TRUE
  } else { out$IQR_reported <- "N" }
  
  if (!is.na(rng_lo)) {
    out$range_reported <- "Y"
    out$range_lower    <- rng_lo
    out$range_upper    <- rng_hi
    any_disp <- TRUE
  } else { out$range_reported <- "N" }
  
  if (!is.na(unspec_lo)) {
    out$unspec_reported <- "Y"
    out$unspec_lower    <- unspec_lo
    out$unspec_upper    <- unspec_hi
    any_disp <- TRUE
  } else { out$unspec_reported <- "N" }
  
  out$dispersion_not_reported <- if (any_disp) "N" else "Y"
  
  # Inherit dispersion type from overall if this subgroup has none
  if (!is.null(inherit_disp) && out$dispersion_not_reported == "Y") {
    if (inherit_disp == "IQR") out$IQR_reported <- "Y"
    if (inherit_disp == "SD")  out$SD_reported  <- "Y"
    out$dispersion_not_reported <- "N"
  }
  
  return(out)
}

# Null coalescing helper
`%||%` <- function(a, b) if (!is.null(a) && !is.na(a)) a else b


# =============================================================================
# PART 2: split_groups()
# =============================================================================
# Returns: overall_reported, subgroups_reported, group_1_name, group_2_name,
#          shared_category, overall_text, group_1_text, group_2_text

split_groups <- function(text) {
  
  empty <- list(
    overall_reported   = NA_character_,
    subgroups_reported = NA_character_,
    group_1_name       = NA_character_,
    group_2_name       = NA_character_,
    shared_category    = NA_character_,
    overall_text       = NA_character_,
    group_1_text       = NA_character_,
    group_2_text       = NA_character_
  )
  
  if (is.na(text) || str_trim(text) == "") return(empty)
  
  text_orig <- normalise_text(text)
  
  # -------------------------------------------------------------------------
  # STEP 1: Strip and save cell header (category prefix + format descriptor)
  # FIX (#10, #15, #20): prevents ">=18 years" becoming group_1_name
  # -------------------------------------------------------------------------
  header_info  <- strip_cell_header(text_orig)
  header_text  <- header_info$header
  working_text <- normalise_text(header_info$remainder)
  
  # Extract shared category from the header or full text
  shared_cat <- extract_category(header_text %||% text_orig)
  
  # -------------------------------------------------------------------------
  # STEP 2: Check for "overall" keyword
  # -------------------------------------------------------------------------
  has_overall_kw <- str_detect(working_text, regex("\\boverall\\b", ignore_case = TRUE))
  overall_text   <- NA_character_
  remaining_text <- working_text
  
  if (has_overall_kw) {
    segs <- str_trim(str_split(working_text, ";")[[1]])
    segs <- segs[segs != ""]
    ov_idx <- which(str_detect(segs, regex("\\boverall\\b", ignore_case = TRUE)))
    if (length(ov_idx) > 0) {
      overall_text   <- segs[ov_idx[1]]
      remaining_segs <- segs[-ov_idx]
      remaining_text <- paste(remaining_segs, collapse = ";")
    }
  }
  
  # -------------------------------------------------------------------------
  # STEP 3: Detect and apply group separator
  # Uses only structural patterns — no hardcoded group names
  # FIX (#92): quoted cells no longer split as groups
  # FIX (#6): value-semicolons not treated as group separators
  # -------------------------------------------------------------------------
  group_segs <- character(0)
  
  # Remove surrounding quotes (FIX #92)
  remaining_clean <- str_remove_all(remaining_text, '^"|"$')
  
  # Pattern A: "keyword - name1 X, name2 Y" (FIX #75)
  # e.g. "median iqr - septic shock 63 (26), noss 60 (30)"
  dash_match <- str_match(remaining_clean, regex(
    "^(?:[a-z\\s]+)?\\s*-\\s*(.+)", ignore_case = TRUE))
  if (!is.na(dash_match[1,1])) {
    after_dash <- str_trim(dash_match[1,2])
    comma_segs <- str_trim(str_split(after_dash, ",\\s*(?=[a-zA-Z])")[[1]])
    comma_segs <- comma_segs[comma_segs != ""]
    if (length(comma_segs) >= 2) group_segs <- comma_segs
  }
  
  # Pattern B: "(n=X) name ... (n=Y) name ..." style — split on "(n=" boundaries
  if (length(group_segs) == 0 &&
      str_count(remaining_clean, regex("\\(n\\s*=\\s*\\d+\\)")) >= 2) {
    group_segs <- str_trim(str_split(remaining_clean,
                                     regex("(?<=\\d)\\s*(?=\\b[a-zA-Z].*?\\(n\\s*=)", ignore_case = TRUE))[[1]])
    group_segs <- group_segs[group_segs != ""]
    if (length(group_segs) < 2) group_segs <- character(0)
  }
  
  # Pattern C: semicolon split (most common)
  if (length(group_segs) == 0 && str_detect(remaining_clean, ";")) {
    if (!is_value_semicolon(remaining_clean)) {
      segs <- str_trim(str_split(remaining_clean, ";")[[1]])
      segs <- segs[segs != ""]
      if (length(segs) >= 2) group_segs <- segs
    }
  }
  
  # Pattern D: "vs" split
  if (length(group_segs) == 0 &&
      str_detect(remaining_clean, regex("\\bvs\\.?\\b", ignore_case = TRUE))) {
    segs <- str_trim(str_split(remaining_clean,
                               regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  
  # Pattern E: sentence split on ". Capital" (e.g. #82, #108, #112)
  if (length(group_segs) == 0 && str_detect(remaining_clean, "\\.\\s+[A-Z]")) {
    segs <- str_trim(str_split(remaining_clean, "(?<=\\.)\\s+(?=[A-Z])")[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  
  # Pattern F: short-label groups "bQ ... bM" (FIX #90)
  if (length(group_segs) == 0) {
    bq_bm <- str_match(remaining_clean, regex(
      "(\\b[a-z]{1,3}Q\\b.+?)(\\b[a-z]{1,3}M\\b.+)", ignore_case = TRUE))
    if (!is.na(bq_bm[1,1])) {
      group_segs <- c(str_trim(bq_bm[1,2]), str_trim(bq_bm[1,3]))
    }
  }
  
  # If no split found and no overall, treat whole original text as overall
  if (length(group_segs) <= 1 && is.na(overall_text)) {
    overall_text <- text_orig
    group_segs   <- character(0)
  }
  
  # -------------------------------------------------------------------------
  # STEP 4: Extract group name from start of each segment
  # Strips leading stat keywords generically — no hardcoded group names
  # -------------------------------------------------------------------------
  extract_group_name <- function(seg) {
    seg <- str_trim(seg)
    
    # Remove leading stat keywords that precede the group name
    # e.g. "Median 63 years (SOC)" -> "(SOC)" -> "SOC"
    # e.g. "Median age: Psittacosis 48" -> "Psittacosis"
    # e.g. "Mean IQR: empiricial antiviral 56.1" -> "empiricial antiviral"
    seg_clean <- str_remove(seg, regex(
      "^(mean|median|average|medican|age|years?|y|IQR|SD)\\s*(age|IQR|SD|years?)?\\s*[:\\s]*",
      ignore_case = TRUE))
    
    # Remove (n=X): notation
    seg_clean <- str_remove(seg_clean, regex("\\(n\\s*=\\s*\\d+[^)]*\\)\\s*:?\\s*", ignore_case = TRUE))
    
    # For "vs" results: try to extract name from brackets first e.g. "(SOC)"
    bracket_name <- str_extract(seg_clean, "\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(bracket_name)) {
      return(str_trim(str_remove_all(bracket_name, "[\\(\\)]")))
    }
    
    # Extract text before first digit or colon
    raw_name <- str_extract(seg_clean, "^[^\\d:]+")
    if (is.na(raw_name)) return(NA_character_)
    
    # Remove trailing noise words
    cleaned <- raw_name
    cleaned <- str_remove(cleaned, regex("\\s*(age\\s*)?(years?|y)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex("\\s*through mechanical ventilation\\.?\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex("\\s*(Their\\s+mean\\s+age[sd]?\\s+were?)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, ":\\s*$")
    cleaned <- str_trim(cleaned)
    
    if (nchar(cleaned) == 0 || str_detect(cleaned, "^[\\d>=<]")) return(NA_character_)
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
    shared_category    = shared_cat,
    overall_text       = overall_text,
    group_1_text       = g1_text,
    group_2_text       = g2_text
  )
}


# =============================================================================
# PART 3: parse_age_cell()
# =============================================================================

parse_age_cell <- function(text) {
  
  groups <- split_groups(text)
  
  overall_data  <- parse_age_block(groups$overall_text)
  group_1_data  <- parse_age_block(
    groups$group_1_text,
    inherit_mean   = overall_data$mean_reported,
    inherit_median = overall_data$median_reported,
    inherit_disp   = if (overall_data$IQR_reported == "Y") "IQR" else
      if (overall_data$SD_reported  == "Y") "SD"  else NULL
  )
  group_2_data  <- parse_age_block(
    groups$group_2_text,
    inherit_mean   = overall_data$mean_reported,
    inherit_median = overall_data$median_reported,
    inherit_disp   = if (overall_data$IQR_reported == "Y") "IQR" else
      if (overall_data$SD_reported  == "Y") "SD"  else NULL
  )
  
  # FIX: shared category propagated to both groups if not already set
  if (!is.na(groups$shared_category)) {
    if (is.na(group_1_data$category)) group_1_data$category <- groups$shared_category
    if (is.na(group_2_data$category)) group_2_data$category <- groups$shared_category
  }
  
  c(
    list(
      overall_reported   = groups$overall_reported,
      subgroups_reported = groups$subgroups_reported,
      group_1_name       = groups$group_1_name,
      group_2_name       = groups$group_2_name
    ),
    setNames(overall_data,  paste0("overall_",  names(overall_data))),
    setNames(group_1_data,  paste0("group_1_",  names(group_1_data))),
    setNames(group_2_data,  paste0("group_2_",  names(group_2_data)))
  )
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
# PART 5: Save to Excel — two sheets
# Sheet 1: NAs left blank (empty cells)
# Sheet 2: NAs replaced with "not reported"
# =============================================================================

wb <- createWorkbook()

# Sheet 1: blank NAs
addWorksheet(wb, "age_cleaned_v4")
writeData(wb, "age_cleaned_v4", my_tibble_clean)

# Sheet 2: "not reported" NAs
my_tibble_nr <- my_tibble_clean %>%
  mutate(across(everything(), ~ ifelse(is.na(.), "not reported", as.character(.))))

addWorksheet(wb, "age_cleaned_v4 (not reported)")
writeData(wb, "age_cleaned_v4 (not reported)", my_tibble_nr)

saveWorkbook(wb, "age_data_cleaned_v4.xlsx", overwrite = TRUE)