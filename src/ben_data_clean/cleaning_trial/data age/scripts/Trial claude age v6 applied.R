# Jesus Is LORD!

# From claude (see separate doc for methodology of using claude)
# Note the line numbers in the contents are incorrect - and thus I do not know whether the contents suggestions are correct

# Addition at line 183 (to load readxl package, and then open file)
# Apply in part 4 and 5 (changing reference terms and file names, etc)

# =============================================================================
# Age Data Parsing Script — Version 6
# =============================================================================
#
# CONTENTS
# --------
# Header: description and change log .......................... 1-108
# Libraries .................................................... 110-112
#
# HELPERS
#   normalise_text() ........................................... 115-125
#   extract_category() ......................................... 128-142
#   is_category_only() ......................................... 145-159
#   is_value_semicolon() ....................................... 162-184
#     Early exit: no semicolon ................................. 163
#     Early exit: Q1;Q3 pattern ............................... 164
#     Split and assess segments ............................... 165-183
#       is_stat_fragment check ................................ 169
#       non_stat_words / has_group_name check ................. 171-177
#     Return: all segments value-like? ........................ 183
#   strip_cell_header() ........................................ 187-244
#     Early exit: no separator ................................ 191-194
#     Split on semicolons ..................................... 196-200
#     Detect standard header .................................. 202-204
#     Detect embedded-groups header ........................... 206-216
#     Build and return header/remainder/avg+disp type ......... 218-243
#   extract_bare_avg() ......................................... 247-260
#     Strip (n=X) notation .................................... 251
#     FIX (#50,56): allow full "years" suffix ................. 253-258
#
# PART 1: parse_age_block() .................................. 263-681
#   Output list initialisation ................................ 270-294
#   Early exit: empty/NA text ................................ 296-303
#   Normalise text and extract category ....................... 305-306
#   Flags (approx, IQR+pm) ................................... 308-311
#   Early exit: category-only cell ........................... 313-321
#
#   AVERAGE EXTRACTION ........................................ 323-466
#     Secondary median removal (#104) ......................... 325-333
#     Mean extraction ......................................... 335-392
#       Special case: "mean +/- SD X ± Y" (#2,8,18) .......... 340-347
#       Pattern A: keyword + optional words + number .......... 349-363
#         FIX (#6): "years?" added to optional words .......... 352
#       Pattern B: "Mean 75" / "Mean 49.5y" .................. 365-370
#       Pattern C: "63 (mean)" ............................... 372-377
#       Pattern D: "Mean (SD) X" ............................. 379-388
#       Pattern E: fallback bare number after keyword ......... 390-402
#     Median extraction ....................................... 404-451
#       Pattern 1: keyword + optional label + number .......... 409-421
#         FIX (#6): "years?" added to optional words .......... 411
#         FIX (#19): bracket label limited to non-digit chars . 412
#         FIX (#92): optional closing bracket skip added ...... 413
#       Pattern 2: number BEFORE keyword (#99) ................ 423-432
#       Pattern 3: "(Median ± SD)" label pattern (#1) ......... 434-444
#         FIX (#1): extract first number, not pre-bracket ..... 436-443
#       Pattern 4: approximate "Median ~65" ................... 446-450
#     Secondary median fallback .............................. 452
#     Average not specified .................................. 454-466
#     Assign average flag columns ............................ 468-478
#     Inheritance: avg type from overall/cell ................. 480-510
#       FIX (#50,56,90,108): value safeguard after inherit .... 505-509
#
#   BUILD post_avg_text ....................................... 512-521
#
#   DISPERSION EXTRACTION ..................................... 523-680
#     Keyword detection flags ................................ 525-528
#     SD extraction ........................................... 530-579
#       Special case: "mean +/- SD X ± Y" (#2,8,18) .......... 533-540
#       "(SD)" label in post_avg_text (#79,84,93) ............. 542-559
#       "SD X" unbracketed ................................... 561-565
#       "± X" / "+/- X" ..................................... 567-572
#       Standalone "+" as SD indicator (#82) .................. 574-578
#     IQR extraction .......................................... 580-626
#       Two-number IQR pattern ............................... 585-592
#       Bracketed pair when keyword elsewhere ................. 593-600
#       Single IQR value ..................................... 602-611
#       Inherited IQR: bracketed single value (#75) .......... 613-625
#     Range extraction ........................................ 628-653
#       "ranged/range from X to Y" pattern ................... 633-636
#       "(range) avg (lo-hi)" pattern ........................ 637-647
#       FIX (#86,89): post_avg_text bracket fallback ......... 648-653
#     Unspecified dispersion ................................. 655-683
#       Bare hyphenated pair ................................. 659-663
#       Single bracketed number -> unspec_disp_value ......... 666-671
#       Bracketed pair (lo, hi) .............................. 672-679
#     Assign dispersion flag columns .......................... 685-732
#       SD ................................................... 687-689
#       IQR (+ capture alongside range/SD) ................... 691-698
#       FIX (#86,89): range assignment isTRUE bug fixed ....... 700-703
#       Unspecified .......................................... 705-711
#       dispersion_not_reported .............................. 713
#       Inherit dispersion type .............................. 715-720
#       FIX (#72,75,90,108): move unspec to IQR if inherited .. 722-732
#     Flag: UQR < LQR ......................................... 734
#     Flag: unparsed numbers ................................. 736-741
#   Return output list ........................................ 743
#
#   Utility functions (%||%, coalesce) ........................ 746-751
#
# PART 2: split_groups() ..................................... 754-886
#   Empty result list ......................................... 759-769
#   Early exit: empty/NA text ................................ 771
#   Normalise and pre-clean text ............................. 773-779
#   Strip and save cell header ............................... 781-784
#   Extract shared category .................................. 786
#   Extract cell-level avg/disp type for inheritance .......... 788-802
#   Detect "overall" keyword segment .......................... 804-818
#   Group splitting patterns ................................. 820-874
#     Pattern A: dash-comma .................................. 822-830
#     Pattern B: (n=X) boundaries ........................... 832-840
#     Pattern C: semicolons .................................. 842-848
#     Pattern D: "vs" ........................................ 850-857
#     Pattern E: ". Capital" sentence split .................. 859-867
#       FIX (#82): merge pronoun-led fragments ............... 862-867
#     Pattern F: short label groups "bQ...bM" ................ 869-874
#   Fallback: treat whole cell as overall ..................... 876-879
#   Extract group name (inner function) ....................... 881-930
#     Strip (n=X) notation ................................... 884
#     Check for "(Name)" at end of segment ................... 886-891
#     FIX (#72): check for plain words at end of segment ..... 893-899
#     Strip leading stat keywords ............................ 901-904
#     Bracketed name at start ............................... 906-910
#     Bracketed name mid-segment ............................ 912-918
#     Text before first digit or colon ....................... 920-931
#   Assign group names and text .............................. 933-937
#   Return result list ........................................ 939-950
#
# PART 3: parse_age_cell() ................................... 953-1000
#   Split cell into group blocks ............................. 956
#   Parse overall block ....................................... 959
#   Determine inheritance values ............................. 961-970
#   Parse group_1 and group_2 blocks ......................... 972-979
#   Propagate shared category ................................ 981-985
#   Combine and return all results ........................... 987-999
#
# PART 4: Apply to tibble .................................... 1002-1015
#   Apply parse_age_cell() to each row ....................... 1005
#   Convert results to tibble ................................ 1007-1009
#   Bind to original tibble .................................. 1011-1012
#
# PART 5: Save to Excel ...................................... 1017-1030
#   Sheet 1: blank NAs ....................................... 1020-1022
#   Sheet 2: "not reported" NAs ............................. 1024-1028
#   Save workbook ............................................ 1030
#
# =============================================================================
# CHANGES VS V5:
#
# EXTRACTION FIXES:
#   #1:  "(Median ± SD)" pattern now extracts FIRST number as median, ± value as SD
#   #6:  "years?" added to optional word list in mean Pattern A and median Pattern 1
#   #19: Optional bracket label in median Pattern 1 now limited to non-digit content
#   #50,56: extract_bare_avg suffix pattern loosened to allow full "years" word
#   #86,89: range assignment block fixed (isTRUE(NA != "Y") bug, same class as prior fix)
#   #86,89: additional range fallback added using post_avg_text bracketed pair
#   #92: median Pattern 1 now skips trailing ")" after optional labels
#
# GROUPING FIXES:
#   #72: group name extraction now checks for plain word(s) at END of segment
#   #72,75,90,108: when inherit_disp="IQR" and dispersion was captured as unspec,
#                  values moved to IQR columns after assignment
#   #82: sentence-split segments starting with common pronouns/articles merged
#        with the preceding segment to prevent false group names
#
# INHERITANCE FIX:
#   #90,108: value set during inheritance now explicitly re-assigned after clearing
#            avg_not_specified, preventing any risk of it being lost
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
normalise_text <- function(text) {
  if (is.na(text)) return(NA_character_)
  text <- str_replace_all(text, "\u2013|\u2014", "-")
  text <- str_replace_all(text, "\u00b1", "±")
  text <- str_replace_all(text, "\u2265", ">=")
  text <- str_replace_all(text, "\u2264", "<=")
  text <- str_replace_all(text, "\u223c", "~")
  str_trim(text)
}

# =============================================================================
# HELPER: extract_category()
# =============================================================================
extract_category <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(NA_character_)
  m <- str_extract(text, regex(paste0(
    "age\\s*\\([^)]*grouped[^)]*\\)|",
    "grouped ages?[^;.]*|",
    "adults?\\s*[>=<]+\\s*\\d+\\s*y(ears?)?|",
    "\\badults?\\b|\\belderly\\b|\\bpaediatric\\b|\\bpediatric\\b|",
    "\\bchild\\b|\\badolescent\\b|\\bgeriatric\\b|",
    "[>=<]+\\s*\\d+\\s*y(ears?)?|",
    "\\d+\\s*-\\s*\\d+\\s*y(ears?)?"
  ), ignore_case = TRUE))
  if (is.na(m)) return(NA_character_)
  str_trim(m)
}

# =============================================================================
# HELPER: is_category_only()
# =============================================================================
is_category_only <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(TRUE)
  t <- str_trim(text)
  stripped <- str_remove(t, regex(paste0(
    "age\\s*\\([^)]*grouped[^)]*\\)|",
    "adults?\\s*[>=<]+\\s*\\d+\\s*y(ears?)?\\s*:?|",
    "\\badults?\\b\\s*:?|\\belderly\\b\\s*:?|",
    "[>=<]+\\s*\\d+\\s*y(ears?)?\\s*:?|",
    "\\d+\\s*-\\s*\\d+\\s*y(ears?)?\\s*:?"
  ), ignore_case = TRUE))
  stripped <- str_trim(stripped)
  nchar(stripped) == 0 || !str_detect(stripped, "\\d")
}

# =============================================================================
# HELPER: is_value_semicolon()
# =============================================================================
is_value_semicolon <- function(text) {
  if (!str_detect(text, ";")) return(FALSE)
  if (str_detect(text, regex("Q1\\s*;\\s*Q3", ignore_case = TRUE))) return(TRUE)
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  stat_kws <- "^\\s*(SD|IQR|Q[123]|mean|median|medican|average|range|years?|y|\\d+\\.?\\d*)\\b"
  value_like <- sapply(segs, function(s) {
    is_stat_fragment <- str_detect(s, regex(stat_kws, ignore_case = TRUE))
    words <- str_extract_all(s, "[a-zA-Z]+")[[1]]
    non_stat_words <- words[!str_detect(words, regex(
      "^(SD|IQR|mean|median|medican|average|range|years?|y|age|and|with|or|the)$",
      ignore_case = TRUE))]
    has_group_name <- length(non_stat_words) >= 2 && str_detect(s, "\\d")
    is_stat_fragment && !has_group_name
  })
  all(value_like)
}

# =============================================================================
# HELPER: strip_cell_header()
# =============================================================================
strip_cell_header <- function(text) {
  result <- list(header = NA_character_, remainder = text,
                 header_avg_type = NA_character_, header_disp_type = NA_character_)
  if (!str_detect(text, ";") && !str_detect(text, "\\bvs\\.?\\b") &&
      !str_detect(text, "\\.\\s+[A-Z]") && !str_detect(text, "\\s-\\s")) {
    return(result)
  }
  segs <- str_trim(str_split(text, ";")[[1]])
  segs <- segs[segs != ""]
  if (length(segs) < 2) return(result)
  first <- segs[1]
  has_category    <- str_detect(first, regex("[>=<]+\\s*\\d+|\\badults?\\b", ignore_case = TRUE))
  has_format_word <- str_detect(first, regex("\\b(mean|median|IQR|SD|years?|y)\\b", ignore_case = TRUE))
  has_age_number  <- str_detect(first, regex("\\b(\\d{2,3}(\\.\\d+)?)\\s*(?:y\\b|[\\(\\[±])", ignore_case = TRUE))
  has_embedded_groups <- FALSE
  if (has_category && has_format_word && has_age_number) {
    stripped_trial <- str_remove(first, regex(
      "^[>=<]?\\s*\\d*\\s*y(ears?)?\\s*:?\\s*(years?|year|y)?\\s*(mean|median|medican|IQR|SD|average)?\\s*(\\[IQR\\]|\\(IQR\\)|\\(SD\\))?\\s*",
      ignore_case = TRUE))
    if (str_detect(stripped_trial, regex("^[A-Za-z].*\\d", ignore_case = FALSE))) {
      has_embedded_groups <- TRUE
    }
  }
  if ((has_category && has_format_word && !has_age_number) || has_embedded_groups) {
    header_avg  <- if (str_detect(first, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))) "median"
    else if (str_detect(first, regex("\\bmean\\b", ignore_case = TRUE))) "mean"
    else NA_character_
    header_disp <- if (str_detect(first, regex("\\bIQR\\b", ignore_case = TRUE))) "IQR"
    else if (str_detect(first, regex("\\bSD\\b", ignore_case = TRUE))) "SD"
    else if (str_detect(first, regex("\\brange\\b", ignore_case = TRUE))) "range"
    else NA_character_
    if (has_embedded_groups) {
      format_stripped <- str_remove(first, regex(
        "^[>=<]?\\s*\\d*\\s*y(ears?)?\\s*:?\\s*(years?|year|y)?\\s*(mean|median|medican|IQR|SD|average)?\\s*(\\[IQR\\]|\\(IQR\\)|\\(SD\\))?\\s*",
        ignore_case = TRUE))
      result$header    <- first
      result$remainder <- paste(c(format_stripped, segs[-1]), collapse = ";")
    } else {
      result$header    <- first
      result$remainder <- paste(segs[-1], collapse = ";")
    }
    result$header_avg_type  <- header_avg
    result$header_disp_type <- header_disp
  }
  result
}

# =============================================================================
# HELPER: extract_bare_avg()
# FIX (#50,56): suffix pattern loosened to allow "years" not just "y"
# =============================================================================
extract_bare_avg <- function(text) {
  if (is.na(text) || str_trim(text) == "") return(NA_real_)
  t <- str_remove_all(text, regex("\\(\\s*n\\s*=\\s*\\d+[^)]*\\)", ignore_case = TRUE))
  t <- str_trim(t)
  # FIX: allow optional "years" or "y" or nothing after the number before any bracket/±/end
  m <- str_extract(t, "^[^\\d]*(\\d+\\.?\\d*)\\s*(?:years?|y)?\\s*(?:[\\(\\[±\\+]|$)")
  if (is.na(m)) return(NA_real_)
  val <- as.numeric(str_extract(m, "\\d+\\.?\\d*"))
  if (!is.na(val) && val >= 5 && val <= 120) return(val)
  NA_real_
}

# =============================================================================
# PART 1: parse_age_block()
# =============================================================================
parse_age_block <- function(text,
                            inherit_mean   = NULL,
                            inherit_median = NULL,
                            inherit_disp   = NULL) {
  
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
    IQR_value              = NA_real_,
    IQR_LQR                = NA_real_,
    IQR_UQR                = NA_real_,
    range_reported         = NA_character_,
    range_lower            = NA_real_,
    range_upper            = NA_real_,
    unspec_reported        = NA_character_,
    unspec_lower           = NA_real_,
    unspec_upper           = NA_real_,
    unspec_disp_value      = NA_real_,
    dispersion_not_reported= NA_character_,
    flag_approx            = NA_character_,
    flag_UQR_lt_LQR        = NA_character_,
    flag_IQR_with_pm       = NA_character_,
    flag_unparsed_numbers  = NA_character_
  )
  
  if (is.na(text) || str_trim(text) == "") {
    out$avg_not_reported        <- "Y"
    out$dispersion_not_reported <- "Y"
    out[c("mean_reported","median_reported","avg_not_specified","SD_reported",
          "IQR_reported","range_reported","unspec_reported")] <- "N"
    out[c("flag_approx","flag_UQR_lt_LQR","flag_IQR_with_pm","flag_unparsed_numbers")] <- "N"
    return(out)
  }
  
  text <- normalise_text(text)
  out$category <- extract_category(text)
  
  out$flag_approx      <- if (str_detect(text, "~")) "Y" else "N"
  out$flag_IQR_with_pm <- if (str_detect(text, regex("\\bIQR\\b", ignore_case=TRUE)) &&
                              str_detect(text, regex("±|\\+/?-|\\+-"))) "Y" else "N"
  
  if (is_category_only(text)) {
    out$avg_not_reported        <- "Y"
    out$dispersion_not_reported <- "Y"
    out[c("mean_reported","median_reported","avg_not_specified","SD_reported",
          "IQR_reported","range_reported","unspec_reported")] <- "N"
    out[c("flag_UQR_lt_LQR","flag_unparsed_numbers")] <- "N"
    return(out)
  }
  
  # -------------------------------------------------------------------------
  # AVERAGE EXTRACTION
  # -------------------------------------------------------------------------
  
  # Secondary median "(median, X)" removal (#104)
  median_secondary <- NA_real_
  secondary_match <- str_extract(text, regex(
    "\\(\\s*median\\s*,?\\s*(\\d+\\.?\\d*)\\s*\\)", ignore_case = TRUE))
  if (!is.na(secondary_match)) {
    median_secondary <- as.numeric(str_extract(secondary_match, "\\d+\\.?\\d*$"))
  }
  text_no_sec <- str_remove(text, regex(
    "\\(\\s*median\\s*,?\\s*\\d+\\.?\\d*\\s*\\)", ignore_case = TRUE))
  
  # --- MEAN ---
  mean_val <- NA_real_
  has_mean_kw <- str_detect(text_no_sec, regex("\\bmean\\b", ignore_case = TRUE))
  
  if (has_mean_kw) {
    # Special case: "mean +/- SD X ± Y" -> X is mean
    pm_sd_match <- str_extract(text_no_sec, regex(
      "\\bmean\\s*\\+/?-\\s*SD\\s+(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(pm_sd_match)) {
      mean_val <- as.numeric(str_extract(pm_sd_match, "\\d+\\.?\\d*$"))
    }
    
    # Pattern A: mean/mean age/mean age was/were/aged/years NUMBER
    # FIX (#6): "years?" added to optional word list
    if (is.na(mean_val)) {
      m1 <- str_extract(text_no_sec, regex(
        "\\bmean\\s*(?:age(?:d|s)?)?\\s*(?:(?:was|were|of|aged|years?|at admission of|,)\\s*)?[~=:]?\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m1)) {
        v <- as.numeric(str_extract(m1, "\\d+\\.?\\d*$"))
        if (!is.na(v) && v >= 5 && v <= 120) mean_val <- v
      }
    }
    
    # Pattern B: "Mean 75", "Mean 49.5y"
    if (is.na(mean_val)) {
      m2 <- str_extract(text_no_sec, regex(
        "\\bmean\\s+(\\d+\\.?\\d*)\\s*y?\\b", ignore_case = TRUE))
      if (!is.na(m2)) mean_val <- as.numeric(str_extract(m2, "\\d+\\.?\\d*"))
    }
    
    # Pattern C: "63 (mean)"
    if (is.na(mean_val)) {
      m3 <- str_extract(text_no_sec, regex(
        "(\\d+\\.?\\d*)\\s*\\(\\s*mean\\s*\\)", ignore_case = TRUE))
      if (!is.na(m3)) mean_val <- as.numeric(str_extract(m3, "^\\d+\\.?\\d*"))
    }
    
    # Pattern D: "Mean (SD) X" or "mean (label) age was X"
    if (is.na(mean_val)) {
      m4 <- str_extract(text_no_sec, regex(
        "\\bmean\\s*\\([^)]*\\)\\s*(?:age(?:d|s)?\\s*(?:was|were|of)?\\s*)?[=:]?\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m4)) {
        v <- as.numeric(str_extract(m4, "\\d+\\.?\\d*$"))
        if (!is.na(v) && v >= 5 && v <= 120) mean_val <- v
      }
    }
    
    # Pattern E: fallback bare number after mean keyword
    if (is.na(mean_val)) {
      stripped <- str_remove(text_no_sec, regex(
        ".*?\\bmean\\b\\s*(?:age(?:d|s)?)?\\s*(?:(?:was|were|of|aged|at admission)\\s*)?[=:,]?\\s*",
        ignore_case = TRUE))
      stripped <- str_remove(stripped, regex("^\\s*\\([^\\d)]*\\)\\s*", ignore_case = TRUE))
      stripped <- str_remove_all(stripped, regex("\\(\\s*n\\s*=\\s*\\d+[^)]*\\)", ignore_case = TRUE))
      m5 <- str_extract(stripped, "^\\s*(\\d+\\.?\\d*)")
      if (!is.na(m5)) {
        v <- as.numeric(str_extract(m5, "\\d+\\.?\\d*"))
        if (!is.na(v) && v >= 5 && v <= 120) mean_val <- v
      }
    }
  }
  
  # --- MEDIAN ---
  median_val <- NA_real_
  has_median_kw <- str_detect(text, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))
  
  if (has_median_kw) {
    # Pattern 1: keyword + optional non-digit bracketed label + optional square bracket
    #            + optional closing bracket skip + optional "age" + number
    # FIX (#6):  "years?" added to optional word list
    # FIX (#19): bracket label limited to (?:\([^\d)]*\))? -- no digits allowed inside
    # FIX (#92): (?:[^(]*\))? added after labels to skip closing ")" before the number
    m1 <- str_extract(text, regex(
      "\\b(?:median|medican)\\s*(?:\\([^\\d)]*\\))?\\s*(?:\\[[^\\]]*\\])?\\s*(?:[^(]*\\))?\\s*(?:age(?:\\s*(?:of|,|was|at))?|years?)?\\s*[~=:,]?\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m1)) {
      val <- as.numeric(str_extract(m1, "\\d+\\.?\\d*$"))
      if (!is.na(val) && val >= 5 && val <= 120) median_val <- val
    }
    
    # Pattern 2: number BEFORE keyword (#99)
    if (is.na(median_val)) {
      m2 <- str_extract(text, regex(
        "(\\d+\\.?\\d*)\\s*(?:\\([^)]*\\))?\\s*(?:median|medican)",
        ignore_case = TRUE))
      if (!is.na(m2)) {
        val <- as.numeric(str_extract(m2, "^\\d+\\.?\\d*"))
        if (!is.na(val) && val >= 5 && val <= 120) median_val <- val
      }
    }
    
    # Pattern 3: "(Median ± SD)" label -- median is FIRST number, SD is after ±
    # FIX (#1): extract first number in cell (not the one immediately before the bracket)
    if (is.na(median_val)) {
      if (str_detect(text, regex("\\(\\s*(?:median|medican)\\s*±", ignore_case = TRUE))) {
        # Extract first plausible age number in the cell
        first_num <- str_extract(text, "\\d+\\.?\\d*")
        if (!is.na(first_num)) {
          val <- as.numeric(first_num)
          if (!is.na(val) && val >= 5 && val <= 120) median_val <- val
        }
      }
    }
    
    # Pattern 4: approximate "Median ~65"
    if (is.na(median_val)) {
      m4 <- str_extract(text, regex("\\bmedian\\b\\s*~\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
      if (!is.na(m4)) median_val <- as.numeric(str_extract(m4, "\\d+\\.?\\d*$"))
    }
  }
  
  if (!is.na(median_secondary) && is.na(median_val)) median_val <- median_secondary
  
  # --- AVG NOT SPECIFIED ---
  avg_ns_val <- NA_real_
  has_avg_kw <- str_detect(text, regex("\\baverage\\b", ignore_case = TRUE))
  
  if (!has_mean_kw && !has_median_kw) {
    avg_ns_val <- extract_bare_avg(text)
    cat_num <- suppressWarnings(as.numeric(str_extract(out$category %||% "", "\\d+")))
    if (!is.na(avg_ns_val) && !is.na(cat_num) && cat_num == avg_ns_val) avg_ns_val <- NA_real_
  } else if (has_avg_kw && !has_mean_kw) {
    m_avg <- str_extract(text, regex(
      "\\baverage\\s*(?:age\\s*(?:was|of)?)?\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m_avg)) avg_ns_val <- as.numeric(str_extract(m_avg, "\\d+\\.?\\d*$"))
  }
  
  # --- ASSIGN AVERAGE FLAGS ---
  out$mean_reported           <- if (!is.na(mean_val) || has_mean_kw) "Y" else "N"
  out$mean_value              <- mean_val
  out$median_reported         <- if (!is.na(median_val) || has_median_kw) "Y" else "N"
  out$median_value            <- median_val
  out$avg_not_specified       <- if (!is.na(avg_ns_val) || has_avg_kw) "Y" else "N"
  out$avg_not_specified_value <- avg_ns_val
  
  out$avg_not_reported <- if (isTRUE(out$mean_reported == "N") &&
                              isTRUE(out$median_reported == "N") &&
                              isTRUE(out$avg_not_specified == "N")) "Y" else "N"
  
  # --- INHERITANCE: avg type ---
  if (isTRUE(out$avg_not_reported == "Y") || isTRUE(out$avg_not_specified == "Y")) {
    bare_val <- if (isTRUE(out$avg_not_reported == "Y")) extract_bare_avg(text) else avg_ns_val
    
    if (!is.null(inherit_median) && isTRUE(inherit_median == "Y") &&
        isTRUE(out$median_reported == "N")) {
      out$median_reported  <- "Y"
      out$median_value     <- bare_val   # set value
      out$avg_not_reported <- "N"
      if (isTRUE(out$avg_not_specified == "Y") && identical(out$avg_not_specified_value, bare_val)) {
        out$avg_not_specified       <- "N"
        out$avg_not_specified_value <- NA_real_
      }
      out$median_value <- bare_val   # FIX (#90,108): re-assign after clearing to prevent loss
      
    } else if (!is.null(inherit_mean) && isTRUE(inherit_mean == "Y") &&
               isTRUE(out$mean_reported == "N")) {
      out$mean_reported    <- "Y"
      out$mean_value       <- bare_val   # set value
      out$avg_not_reported <- "N"
      if (isTRUE(out$avg_not_specified == "Y") && identical(out$avg_not_specified_value, bare_val)) {
        out$avg_not_specified       <- "N"
        out$avg_not_specified_value <- NA_real_
      }
      out$mean_value <- bare_val   # FIX (#90,108): re-assign after clearing to prevent loss
    }
  }
  
  # -------------------------------------------------------------------------
  # BUILD post_avg_text
  # -------------------------------------------------------------------------
  avg_val_used <- coalesce(mean_val, median_val, avg_ns_val,
                           out$mean_value, out$median_value, out$avg_not_specified_value)
  post_avg_text <- text
  if (!is.na(avg_val_used)) {
    pos <- str_locate(text, fixed(as.character(avg_val_used)))[1, "end"]
    if (!is.na(pos)) post_avg_text <- substr(text, pos + 1, nchar(text))
  }
  
  # -------------------------------------------------------------------------
  # DISPERSION EXTRACTION
  # -------------------------------------------------------------------------
  has_iqr   <- str_detect(text, regex("\\bIQR\\b|Q1.*Q3|25/?75%?\\s*IQR", ignore_case = TRUE))
  has_range <- str_detect(text, regex("\\brange[d]?\\b", ignore_case = TRUE))
  has_sd_kw <- str_detect(text, regex("\\bSD\\b|±|\\+\\s*/?\\s*-|\\+-", ignore_case = FALSE)) ||
    str_detect(text, regex("\\bSD\\b", ignore_case = TRUE))
  
  # --- SD ---
  sd_val <- NA_real_
  
  # "mean +/- SD X ± Y" -> Y is SD
  pm_sd_disp <- str_extract(text, regex(
    "\\bmean\\s*\\+/?-\\s*SD\\s+\\d+\\.?\\d*\\s*(?:±|\\+/?-)\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE))
  if (!is.na(pm_sd_disp)) {
    nums <- as.numeric(str_extract_all(pm_sd_disp, "\\d+\\.?\\d*")[[1]])
    if (length(nums) >= 2) sd_val <- nums[length(nums)]
  }
  
  # "(SD)" bracketed label -> search post_avg_text
  if (is.na(sd_val) && str_detect(text, regex("\\bSD\\b", ignore_case = TRUE))) {
    m_sd_post <- str_extract(post_avg_text, regex(
      "\\(\\s*SD\\s*\\)\\s*[=:]?\\s*(\\d+\\.?\\d*)|\\(\\s*SD\\s+(\\d+\\.?\\d*)\\s*\\)",
      ignore_case = TRUE))
    if (!is.na(m_sd_post)) {
      nums <- str_extract_all(m_sd_post, "\\d+\\.?\\d*")[[1]]
      if (length(nums) > 0) sd_val <- as.numeric(nums[length(nums)])
    }
    if (is.na(sd_val) && !has_iqr) {
      m_brk <- str_extract(post_avg_text, "\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
      if (!is.na(m_brk)) {
        val <- as.numeric(str_extract(m_brk, "\\d+\\.?\\d*"))
        if (!is.na(val) && !str_detect(m_brk, "[,;\\-]")) sd_val <- val
      }
    }
  }
  
  if (is.na(sd_val) && !has_iqr) {
    m_sd1 <- str_extract(post_avg_text, regex("\\bSD\\b\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
    if (!is.na(m_sd1)) sd_val <- as.numeric(str_extract(m_sd1, "\\d+\\.?\\d*$"))
  }
  
  if (is.na(sd_val) && !has_iqr) {
    m_sd2 <- str_extract(post_avg_text, regex(
      "(?:±|\\+\\s*/?\\s*-|\\+-|plus\\s*/?\\s*minus)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd2)) sd_val <- as.numeric(str_extract(m_sd2, "\\d+\\.?\\d*$"))
  }
  
  # Standalone "+" as SD indicator (#82)
  if (is.na(sd_val) && !has_iqr) {
    m_sd3 <- str_extract(post_avg_text, regex("(?<![+/\\-])\\+(?![+/\\-])\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd3)) sd_val <- as.numeric(str_extract(m_sd3, "\\d+\\.?\\d*$"))
  }
  
  # FIX (#1): "(Median ± SD)" label -> extract SD from ± BEFORE the bracket
  if (is.na(sd_val) && str_detect(text, regex("\\(\\s*(?:median|medican)\\s*±", ignore_case = TRUE))) {
    m_med_sd <- str_extract(text, regex("(\\d+\\.?\\d*)\\s*\\(\\s*(?:median|medican)\\s*±", ignore_case = TRUE))
    if (!is.na(m_med_sd)) {
      # The number before the bracket is the SD (e.g. 14.7 in "42.3 ± 14.7 (Median ± SD)")
      sd_val <- as.numeric(str_extract(m_med_sd, "^\\d+\\.?\\d*"))
    }
  }
  
  # --- IQR ---
  iqr_val <- NA_real_
  iqr_lo  <- NA_real_
  iqr_hi  <- NA_real_
  
  if (has_iqr) {
    m_iqr <- str_extract(text, regex(
      "(?:IQR|Q1\\s*[;,]?\\s*Q3|25/?75%?\\s*IQR)\\s*[:\\(\\[]?\\s*(\\d+\\.?\\d*)\\s*(?:[-,;]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (!is.na(m_iqr)) {
      nums <- as.numeric(str_extract_all(m_iqr, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
    }
    if (is.na(iqr_lo)) {
      m_iqr2 <- str_extract(post_avg_text,
                            "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[,;\\-]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_iqr2)) {
        nums <- as.numeric(str_extract_all(m_iqr2, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
      }
    }
    if (is.na(iqr_lo)) {
      m_iqr3 <- str_extract(post_avg_text, regex(
        "IQR\\s*:?\\s*(\\d+\\.?\\d*)|(?:±|\\+\\s*/?\\s*-|\\+-)\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m_iqr3)) {
        nums <- as.numeric(str_extract_all(m_iqr3, "\\d+\\.?\\d*")[[1]])
        if (length(nums) > 0) iqr_val <- nums[length(nums)]
      }
    }
    if (is.na(iqr_lo) && is.na(iqr_val) && !is.null(inherit_disp) &&
        isTRUE(inherit_disp == "IQR")) {
      m_iqr4 <- str_extract(post_avg_text, "\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
      if (!is.na(m_iqr4)) {
        val <- as.numeric(str_extract(m_iqr4, "\\d+\\.?\\d*"))
        if (!is.na(val) && !str_detect(m_iqr4, "[,;\\-]")) iqr_val <- val
      }
    }
  }
  
  # --- RANGE ---
  rng_lo <- NA_real_; rng_hi <- NA_real_
  if (has_range) {
    m_rng <- str_extract(text, regex(
      "range[d]?\\s*(?:from)?\\s*(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)",
      ignore_case = TRUE))
    if (is.na(m_rng)) {
      m_rng2 <- str_extract(text, regex(
        "\\(\\s*range\\s*\\)[^\\d]*(\\d+\\.?\\d*)[^\\d]+(\\d+\\.?\\d*)\\s*(?:to|-|,)\\s*(\\d+\\.?\\d*)",
        ignore_case = TRUE))
      if (!is.na(m_rng2)) {
        nums <- as.numeric(str_extract_all(m_rng2, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 3) { rng_lo <- nums[2]; rng_hi <- nums[3] }
      } else {
        # FIX (#86,89): fallback — look for bracketed pair in post_avg_text
        # when "range" keyword present but numbers follow the average
        m_rng3 <- str_extract(post_avg_text,
                              "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
        if (!is.na(m_rng3)) {
          nums <- as.numeric(str_extract_all(m_rng3, "\\d+\\.?\\d*")[[1]])
          if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
        }
      }
    } else {
      nums <- as.numeric(str_extract_all(m_rng, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
    }
  }
  
  # --- UNSPECIFIED ---
  unspec_lo  <- NA_real_; unspec_hi <- NA_real_
  unspec_val <- NA_real_
  
  m_bare <- str_extract(str_trim(text), "^\\s*(\\d+\\.?\\d*)\\s*-\\s*(\\d+\\.?\\d*)\\s*$")
  if (!is.na(m_bare)) {
    nums <- as.numeric(str_extract_all(m_bare, "\\d+\\.?\\d*")[[1]])
    if (length(nums) == 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
  }
  
  has_any_kw <- str_detect(text, regex(
    "\\bSD\\b|\\bIQR\\b|\\brange\\b|\\bQ1\\b|±|\\+/?-", ignore_case = TRUE))
  
  if (is.na(unspec_lo) && !has_any_kw) {
    m_single <- str_extract(post_avg_text, "^\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
    if (!is.na(m_single) && !str_detect(m_single, "[,;\\-]")) {
      unspec_val <- as.numeric(str_extract(m_single, "\\d+\\.?\\d*"))
    } else {
      m_unspec <- str_extract(post_avg_text,
                              "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_unspec)) {
        nums <- as.numeric(str_extract_all(m_unspec, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
      }
    }
  }
  
  # --- ASSIGN DISPERSION FLAGS ---
  any_disp <- FALSE
  
  if (!is.na(sd_val)) {
    out$SD_reported <- "Y"; out$SD_value <- sd_val; any_disp <- TRUE
  } else { out$SD_reported <- "N" }
  
  if (has_iqr) {
    out$IQR_reported <- "Y"; any_disp <- TRUE
    out$IQR_LQR <- iqr_lo; out$IQR_UQR <- iqr_hi; out$IQR_value <- iqr_val
    if (!is.na(rng_lo)) {
      out$range_reported <- "Y"; out$range_lower <- rng_lo; out$range_upper <- rng_hi
    }
    if (!is.na(sd_val) && is.na(out$SD_value)) out$SD_value <- sd_val
  } else { out$IQR_reported <- "N" }
  
  # FIX (#86,89): was isTRUE(out$range_reported != "Y") which returns FALSE when NA
  # Now uses !isTRUE(out$range_reported == "Y") which correctly returns TRUE when NA
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
  
  # Inherit dispersion type when none found
  if (isTRUE(out$dispersion_not_reported == "Y") && !is.null(inherit_disp)) {
    if (isTRUE(inherit_disp == "IQR"))   { out$IQR_reported   <- "Y"; out$dispersion_not_reported <- "N" }
    if (isTRUE(inherit_disp == "SD"))    { out$SD_reported    <- "Y"; out$dispersion_not_reported <- "N" }
    if (isTRUE(inherit_disp == "range")) { out$range_reported <- "Y"; out$dispersion_not_reported <- "N" }
  }
  
  # FIX (#72,75,90,108): when inherit_disp="IQR" and dispersion was captured as unspec
  # (because the IQR keyword was in the header/overall but not in this subgroup text),
  # move unspec values into IQR columns
  if (!is.null(inherit_disp) && isTRUE(inherit_disp == "IQR") &&
      isTRUE(out$unspec_reported == "Y")) {
    out$IQR_reported <- "Y"
    out$dispersion_not_reported <- "N"
    # Move unspec pair to IQR bounds
    if (!is.na(out$unspec_lower)) {
      out$IQR_LQR    <- out$unspec_lower
      out$IQR_UQR    <- out$unspec_upper
      out$unspec_lower <- NA_real_
      out$unspec_upper <- NA_real_
    }
    # Move single unspec value to IQR_value
    if (!is.na(out$unspec_disp_value)) {
      out$IQR_value        <- out$unspec_disp_value
      out$unspec_disp_value <- NA_real_
    }
    out$unspec_reported <- "N"
  }
  
  out$flag_UQR_lt_LQR <- if (!is.na(iqr_lo) && !is.na(iqr_hi) && iqr_hi < iqr_lo) "Y" else "N"
  
  all_nums_in_cell <- as.numeric(str_extract_all(text, "\\d+\\.?\\d*")[[1]])
  all_nums_in_cell <- all_nums_in_cell[!is.na(all_nums_in_cell)]
  out_nums <- unlist(out[sapply(out, is.numeric)])
  out_nums <- out_nums[!is.na(out_nums)]
  out$flag_unparsed_numbers <- if (length(all_nums_in_cell) > (length(out_nums) + 1)) "Y" else "N"
  
  return(out)
}

`%||%` <- function(a, b) if (!is.null(a) && length(a) > 0 && !is.na(a[1])) a else b
coalesce <- function(...) {
  args <- list(...)
  for (a in args) { if (!is.null(a) && length(a) > 0 && !is.na(a[1])) return(a) }
  NA_real_
}

# =============================================================================
# PART 2: split_groups()
# =============================================================================
split_groups <- function(text) {
  
  empty <- list(
    overall_reported   = NA_character_,
    subgroups_reported = NA_character_,
    group_1_name       = NA_character_,
    group_2_name       = NA_character_,
    shared_category    = NA_character_,
    overall_text       = NA_character_,
    group_1_text       = NA_character_,
    group_2_text       = NA_character_,
    cell_avg_type      = NA_character_,
    cell_disp_type     = NA_character_
  )
  
  if (is.na(text) || str_trim(text) == "") return(empty)
  
  text_orig <- normalise_text(text)
  text_work <- str_remove_all(text_orig, '^"|"$')
  text_work <- str_remove(text_work, regex("^Age\\s*,?\\s*", ignore_case = TRUE))
  text_work <- normalise_text(text_work)
  
  header_info  <- strip_cell_header(text_work)
  header_text  <- header_info$header
  working_text <- normalise_text(header_info$remainder)
  
  shared_cat     <- extract_category(header_text %||% text_orig)
  cell_avg_type  <- header_info$header_avg_type
  cell_disp_type <- header_info$header_disp_type
  
  if (is.na(cell_avg_type)) {
    cell_avg_type <- if (str_detect(text_work, regex("\\bmedian\\b|\\bmedican\\b", ignore_case=TRUE))) "median"
    else if (str_detect(text_work, regex("\\bmean\\b", ignore_case=TRUE))) "mean"
    else NA_character_
  }
  if (is.na(cell_disp_type)) {
    cell_disp_type <- if (str_detect(text_work, regex("\\bIQR\\b", ignore_case=TRUE))) "IQR"
    else if (str_detect(text_work, regex("\\bSD\\b", ignore_case=TRUE))) "SD"
    else if (str_detect(text_work, regex("\\brange\\b", ignore_case=TRUE))) "range"
    else NA_character_
  }
  
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
  
  group_segs <- character(0)
  
  # Pattern A: dash-comma
  dash_match <- str_match(remaining_text, regex("^(?:[a-z\\s]+)?\\s*-\\s*(.+)", ignore_case = TRUE))
  if (!is.na(dash_match[1,1])) {
    after_dash <- str_trim(dash_match[1,2])
    comma_segs <- str_trim(str_split(after_dash, ",\\s*(?=[a-zA-Z])")[[1]])
    comma_segs <- comma_segs[comma_segs != ""]
    if (length(comma_segs) >= 2) group_segs <- comma_segs
  }
  
  # Pattern B: (n=X) boundaries
  if (length(group_segs) == 0 &&
      str_count(remaining_text, regex("\\(n\\s*=\\s*\\d+\\)")) >= 2) {
    parts <- str_trim(unlist(str_split(remaining_text,
                                       regex("(?<=\\d\\)\\s{0,3})(?=[a-zA-Z])"))))
    parts <- parts[parts != "" & str_detect(parts, "\\d")]
    if (length(parts) >= 2) group_segs <- parts
  }
  
  # Pattern C: semicolons
  if (length(group_segs) == 0 && str_detect(remaining_text, ";")) {
    if (!is_value_semicolon(remaining_text)) {
      segs <- str_trim(str_split(remaining_text, ";")[[1]])
      segs <- segs[segs != ""]
      if (length(segs) >= 2) group_segs <- segs
    }
  }
  
  # Pattern D: "vs"
  if (length(group_segs) == 0 &&
      str_detect(remaining_text, regex("\\bvs\\.?\\b", ignore_case = TRUE))) {
    segs <- str_trim(str_split(remaining_text,
                               regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
    segs <- segs[segs != ""]
    if (length(segs) >= 2) group_segs <- segs
  }
  
  # Pattern E: ". Capital" sentence split
  # FIX (#82): after splitting, merge any segment that starts with a common
  # pronoun or article back into the preceding segment
  if (length(group_segs) == 0 && str_detect(remaining_text, "\\.\\s+[A-Z]")) {
    segs <- str_trim(str_split(remaining_text, "(?<=\\.)\\s+(?=[A-Z])")[[1]])
    segs <- segs[segs != ""]
    # Merge pronoun/article-led fragments with the previous segment
    pronouns <- regex("^(Their|His|Her|Its|The|A|An|This|These|Those)\\b", ignore_case = FALSE)
    merged <- character(0)
    for (s in segs) {
      if (length(merged) > 0 && str_detect(s, pronouns)) {
        merged[length(merged)] <- paste(merged[length(merged)], s)
      } else {
        merged <- c(merged, s)
      }
    }
    if (length(merged) >= 2) group_segs <- merged
  }
  
  # Pattern F: short label groups "bQ...bM"
  if (length(group_segs) == 0) {
    bq_bm <- str_match(remaining_text, regex(
      "(\\b[a-z]{1,3}Q\\b.+?)(\\b[a-z]{1,3}M\\b.+)", ignore_case = TRUE))
    if (!is.na(bq_bm[1,1])) {
      group_segs <- c(str_trim(bq_bm[1,2]), str_trim(bq_bm[1,3]))
    }
  }
  
  if (length(group_segs) <= 1 && is.na(overall_text)) {
    overall_text <- text_orig
    group_segs   <- character(0)
  }
  
  # --- Extract group name ---
  extract_group_name <- function(seg) {
    seg <- str_trim(seg)
    seg_clean <- str_remove(seg, regex("\\(n\\s*=\\s*\\d+[^)]*\\)\\s*:?\\s*", ignore_case = TRUE))
    
    # Check for "(Name)" at END of segment
    end_bracket <- str_extract(seg_clean, regex("\\(\\s*([A-Z][A-Za-z0-9\\-\\s]+)\\s*\\)\\s*$"))
    if (!is.na(end_bracket)) {
      name <- str_trim(str_remove_all(end_bracket, "[\\(\\)]"))
      if (nchar(name) >= 2) return(name)
    }
    
    # FIX (#72): check for plain word(s) at END of segment after numbers and brackets
    # e.g. "48 (36-55) indigenous" -> "indigenous"
    end_words <- str_extract(seg_clean, regex(
      "(?:[\\d\\s\\(\\)\\[\\].,;±+-]+)([a-zA-Z][a-zA-Z\\s-]+)$"))
    if (!is.na(end_words)) {
      name <- str_trim(str_extract(end_words, "[a-zA-Z][a-zA-Z\\s-]+$"))
      # Only use if it's not a stat keyword and is reasonably long
      if (!is.na(name) && nchar(name) >= 3 &&
          !str_detect(name, regex("^(years?|y|SD|IQR|range|mean|median|average)$", ignore_case = TRUE))) {
        return(str_trim(name))
      }
    }
    
    # Strip leading stat keywords
    seg_clean <- str_remove(seg_clean, regex(
      "^(mean|median|medican|average|age|years?|y|IQR|SD)\\s*(age|IQR|SD|years?)?\\s*[:\\s]*",
      ignore_case = TRUE))
    
    # Bracketed name at start
    bracket_name <- str_extract(seg_clean, "^\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(bracket_name)) {
      return(str_trim(str_remove_all(bracket_name, "[\\(\\)]")))
    }
    
    # Bracketed name mid-segment
    bracket_mid <- str_extract(seg_clean, "\\(\\s*([A-Z][A-Za-z0-9\\-]+)\\s*\\)")
    if (!is.na(bracket_mid)) {
      name <- str_trim(str_remove_all(bracket_mid, "[\\(\\)]"))
      if (nchar(name) >= 2 && !str_detect(name, regex(
        "^(IQR|SD|range|mean|median)$", ignore_case = TRUE)))
        return(name)
    }
    
    raw_name <- str_extract(seg_clean, "^[^\\d:]+")
    if (is.na(raw_name)) return(NA_character_)
    
    cleaned <- raw_name
    cleaned <- str_remove(cleaned, regex("\\s*(age\\s*)?(years?|y)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex(
      "\\s*\\(\\s*(range|IQR|SD|mean|median)\\s*\\)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex(
      "\\s*through mechanical ventilation\\.?\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, regex(
      "\\s*(mean|median|IQR|SD|average)\\s*:?\\s*$", ignore_case = TRUE))
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
    group_2_text       = g2_text,
    cell_avg_type      = cell_avg_type,
    cell_disp_type     = cell_disp_type
  )
}

# =============================================================================
# PART 3: parse_age_cell()
# =============================================================================
parse_age_cell <- function(text) {
  
  groups <- split_groups(text)
  overall_data <- parse_age_block(groups$overall_text)
  
  inh_mean   <- if (isTRUE(overall_data$mean_reported   == "Y")) "Y"
  else if (!is.na(groups$cell_avg_type) && groups$cell_avg_type == "mean")   "Y"
  else "N"
  inh_median <- if (isTRUE(overall_data$median_reported == "Y")) "Y"
  else if (!is.na(groups$cell_avg_type) && groups$cell_avg_type == "median") "Y"
  else "N"
  inh_disp   <- if (isTRUE(overall_data$IQR_reported == "Y"))   "IQR"
  else if (isTRUE(overall_data$SD_reported  == "Y")) "SD"
  else if (!is.na(groups$cell_disp_type)) groups$cell_disp_type
  else NULL
  
  group_1_data <- parse_age_block(groups$group_1_text,
                                  inherit_mean   = inh_mean,
                                  inherit_median = inh_median,
                                  inherit_disp   = inh_disp)
  group_2_data <- parse_age_block(groups$group_2_text,
                                  inherit_mean   = inh_mean,
                                  inherit_median = inh_median,
                                  inherit_disp   = inh_disp)
  
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
# Replace "my_tibble" and "age_column" with your actual names.
# =============================================================================
parsed <- lapply(age$age, parse_age_cell)

parsed_tibble <- bind_rows(
  lapply(parsed, function(x) as.data.frame(x, stringsAsFactors = FALSE))
)

my_tibble_clean <- bind_cols(age, parsed_tibble)
glimpse(my_tibble_clean)

# =============================================================================
# PART 5: Save to Excel — two sheets
# =============================================================================
wb <- createWorkbook()

addWorksheet(wb, "age_clean_v6")
writeData(wb, "age_clean_v6", my_tibble_clean)

my_tibble_nr <- my_tibble_clean %>%
  mutate(across(everything(), ~ ifelse(is.na(.), "NR", as.character(.))))

addWorksheet(wb, "age_clean_v6 (not reported)")
writeData(wb, "age_clean_v6 (not reported)", my_tibble_nr)

saveWorkbook(wb, "age_data_cleaned_v6.xlsx", overwrite = TRUE)