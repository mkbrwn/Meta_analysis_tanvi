# =============================================================================
# Proportion Data Cleaning Script
# =============================================================================
#
# CONTENTS
# --------
# Header: description ............................................. 1-70
# Libraries ....................................................... 72-74
#
# CONFIGURATION ................................................... 76-120
#   File paths .................................................... 77-79
#   Column name constants ......................................... 81-100
#   Proportion columns to process ................................ 102-120
#
# HELPERS ......................................................... 122-400
#   normalise_text() .............................................. 123-132
#   clean_n_total() ............................................... 135-160
#     Cleans column S (number of participants) .................... 136-159
#   detect_subgroups() ............................................ 162-198
#     Detects subgroups from columns T, BB, BC .................... 163-197
#   split_outside_parens() ........................................ 200-221
#     Splits text on delimiter outside brackets ................... 201-220
#   extract_leading_label() ....................................... 223-234
#     Extracts text label from start of a segment ................. 224-233
#   fix_typo_fractions() .......................................... 236-244
#     Fixes X/X/Y and X/Y/Y typos to X/Y ......................... 237-243
#   parse_single_block() .......................................... 246-342
#     Parses one text block into num/prop/pop/pct ................. 247-341
#       Pattern: N/M (pct%) ....................................... 261-268
#       Pattern: N/M (no percent) ................................. 270-276
#       Pattern: N (pct%) ......................................... 278-285
#       Pattern: N% ............................................... 287-293
#       Pattern: N=X ............................................... 295-302
#       Pattern: direct proportion (0-1) .......................... 304-312
#       Pattern: first integer + second decimal ................... 314-321
#       Pattern: first >100 + second <=100 ........................ 323-330
#       Pattern: standalone integer ............................... 332-340
#   split_into_blocks() ........................................... 344-382
#     Splits cell into labelled subgroup blocks ................... 345-381
#   parse_proportion_cell() ....................................... 384-452
#     Master function: parses full cell ........................... 385-451
#   safe_eq() ..................................................... 454-465
#     Approximate equality check for validation ................... 455-464
#   flag_duplication() ............................................ 467-485
#     Flags numbers appearing in other proportion columns ......... 468-484
#   build_study_reference() ....................................... 487-496
#     Builds "Surname Year" reference string ...................... 488-495
#
# PART 1: Clean column S (number of participants) ................. 498-508
#
# PART 2: Build row-level number sets for duplication check ....... 510-522
#
# PART 3: Process each proportion column .......................... 524-650
#   Loop over PROPORTION_COLS ..................................... 525-649
#     Build combined source series ............................... 527-532
#     Per-row processing ......................................... 534-645
#       Parse cell ............................................... 536-537
#       Assign overall and subgroup blocks ....................... 539-572
#       Compute check values (manual_prop, manual_num, pop_equal) 574-596
#       Compute flags (unused numbers, duplication) .............. 598-618
#       Assemble output row ..................................... 620-648
#
# PART 4: Write output workbooks .................................. 652-700
#   Sheet 1: blank NAs ........................................... 660-675
#   Sheet 2: NR NAs .............................................. 677-695
#   Save both workbooks ........................................... 697-700
#
# =============================================================================
# COLUMN MAPPING (Excel letter -> column header in merged_sheets_updated_1.xlsx)
#   D  = study number
#   F  = authors
#   I  = year published
#   S  = number of participants  [used for validation checks]
#   T  = groups if applicable (e.g. if case-control, cohort)
#   BB = intervention group
#   BC = control group
#   X  = icu mortality
#   Y  = 28d mortality
#   Z  = 30d mortality
#   AA = 60d mortality
#   AB = 90d mortality
#   AC = hospital mortality
#   AF = survival to hospital discharge
#   AG = mechnical ventilation (proportion requiring)
#   AH = niv (proportion requiring)
#   AI = ecmo (proportion requiring)
#   AJ = rrt (proportion requiring)
#   AK = vasopressor / inotropic support (proportion requiring)
#   AL = septic shock (proportion with)
#   AM = respiratory failure (proportion with)
#   AN = ards (proportion with)
#   AP = bacterial (proportion)
#   AQ = viral (proportion)
#   AT = sex           |
#   BD = males         | -> combined sex sheet
#   BV = male          |
#   AU = copd
#   AV = smokers       |
#   BW = smoking       | -> combined smoking sheet
#   AW = diabetes
#   BG = 6 months mortality
#
# REPLACE input_path and output_path* below with your actual file paths.
# =============================================================================

library(dplyr)
library(stringr)
library(openxlsx)
library(tidyr)

# =============================================================================
# CONFIGURATION
# =============================================================================

input_path      <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/merged_sheets_updated_1.xlsx"
output_path     <- "proportion_cleaned.xlsx"
output_path_nr  <- "proportion_cleaned_NR.xlsx"

# Column name constants (matching headers in merged_sheets_updated_1.xlsx)
COL_STUDY_NUM  <- "study number"
COL_AUTHORS    <- "authors"
COL_YEAR       <- "year published"
COL_N_TOTAL    <- "number of participants"
COL_GROUPS     <- "groups if applicable (e.g. if case-control, cohort)"
COL_GROUP1     <- "intervention group"      # BB
COL_GROUP2     <- "control group"           # BC
COL_SEX_AT     <- "sex"                     # AT
COL_SEX_BD     <- "males"                   # BD
COL_SEX_BV     <- "male"                    # BV
COL_SMOKE_AV   <- "smokers"                 # AV
COL_SMOKE_BW   <- "smoking"                 # BW

# Each entry: list(sheet_name, character vector of source column names)
# For combined columns (sex, smoking), all source columns are listed;
# the first non-NA value across them is used.
PROPORTION_COLS <- list(
  list("S_n_participants",  c(COL_N_TOTAL)),
  list("X_icu_mortality",   c("icu mortality")),
  list("Y_28d_mortality",   c("28d mortality")),
  list("Z_30d_mortality",   c("30d mortality")),
  list("AA_60d_mortality",  c("60d mortality")),
  list("AB_90d_mortality",  c("90d mortality")),
  list("AC_hospital_mort",  c("hospital mortality")),
  list("AF_survival_disch", c("survival to hospital discharge")),
  list("AG_mech_vent",      c("mechnical ventilation (proportion requiring)")),
  list("AH_niv",            c("niv (proportion requiring)")),
  list("AI_ecmo",           c("ecmo (proportion requiring)")),
  list("AJ_rrt",            c("rrt (proportion requiring)")),
  list("AK_vasopressor",    c("vasopressor / inotropic support (proportion requiring)")),
  list("AL_septic_shock",   c("septic shock (proportion with)")),
  list("AM_resp_failure",   c("respiratory failure (proportion with)")),
  list("AN_ards",           c("ards (proportion with)")),
  list("AP_bacterial",      c("bacterial (proportion)")),
  list("AQ_viral",          c("viral (proportion)")),
  list("AT_BD_BV_sex",      c(COL_SEX_AT, COL_SEX_BD, COL_SEX_BV)),
  list("AU_copd",           c("copd")),
  list("AV_BW_smoking",     c(COL_SMOKE_AV, COL_SMOKE_BW)),
  list("AW_diabetes",       c("diabetes")),
  list("BG_6month_mort",    c("6 months mortality"))
)

# =============================================================================
# HELPERS
# =============================================================================

# normalise_text: standardise special characters
normalise_text <- function(x) {
  if (is.na(x) || str_trim(x) == "") return(NA_character_)
  x <- str_replace_all(x, "\u2009", " ")   # thin space
  x <- str_replace_all(x, "\u00a0", " ")   # non-breaking space
  x <- str_replace_all(x, "\u00b1", "±")
  x <- str_trim(x)
  x
}

# ---------------------------------------------------------------------------
# clean_n_total: extract primary N from column S
# Handles: "86", "n = 575 (development n=455...)", "total n=204", etc.
# Returns integer or NA
# ---------------------------------------------------------------------------
clean_n_total <- function(x) {
  if (is.na(x) || str_trim(x) == "") return(NA_integer_)
  s <- str_trim(as.character(x))
  
  # Direct integer
  if (grepl("^\\d+$", s)) return(as.integer(s))
  
  # "total n=X"
  m <- str_match(s, regex("total\\s*n\\s*[=:]\\s*(\\d+)", ignore_case = TRUE))
  if (!is.na(m[1,1])) return(as.integer(m[1,2]))
  
  # "n=X" or "n = X"
  m <- str_match(s, regex("\\bn\\s*[=:]\\s*(\\d+)", ignore_case = TRUE))
  if (!is.na(m[1,1])) return(as.integer(m[1,2]))
  
  # Leading number
  m <- str_match(s, "^\\s*(\\d+)")
  if (!is.na(m[1,1])) return(as.integer(m[1,2]))
  
  NA_integer_
}

# ---------------------------------------------------------------------------
# detect_subgroups: uses columns T, BB, BC to detect subgroups
# Returns list(has_subgroups=TRUE/FALSE, n_subgroups=integer or NA)
# ---------------------------------------------------------------------------
CONTRAST_PAT <- regex(
  "\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|\\bv\\b|\\band\\b|;|/|\\bgroup\\b",
  ignore_case = TRUE)

detect_subgroups <- function(groups_val, g1_val, g2_val) {
  g  <- if (isTRUE(!is.na(groups_val)) && length(groups_val) > 0) str_trim(as.character(groups_val)) else ""
  g1 <- if (isTRUE(!is.na(g1_val))    && length(g1_val)    > 0) str_trim(as.character(g1_val))     else ""
  g2 <- if (isTRUE(!is.na(g2_val))    && length(g2_val)    > 0) str_trim(as.character(g2_val))     else ""
  
  # "single" without contrast => no subgroups
  if (str_detect(g, regex("\\bsingle\\b", ignore_case = TRUE)) &&
      !str_detect(g, CONTRAST_PAT)) {
    return(list(has_subgroups = FALSE, n_subgroups = 0L))
  }
  
  # BB and BC both populated => 2 subgroups
  if (nchar(g1) > 0 && nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L))
  if (nchar(g1) > 0 || nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L))
  
  # Count parts split by vs/;
  parts <- str_trim(str_split(g, regex("\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|;",
                                       ignore_case = TRUE))[[1]])
  parts <- parts[nchar(parts) > 0]
  if (length(parts) >= 2)
    return(list(has_subgroups = TRUE, n_subgroups = length(parts)))
  
  if (str_detect(g, CONTRAST_PAT))
    return(list(has_subgroups = TRUE, n_subgroups = NA_integer_))
  
  list(has_subgroups = FALSE, n_subgroups = 0L)
}

# ---------------------------------------------------------------------------
# split_outside_parens: split on delimiter only outside brackets/parens
# ---------------------------------------------------------------------------
split_outside_parens <- function(text, delimiter = ";") {
  parts   <- character(0)
  depth   <- 0L
  current <- character(0)
  chars   <- strsplit(text, "")[[1]]
  for (ch in chars) {
    if (ch %in% c("(", "["))  depth <- depth + 1L
    else if (ch %in% c(")", "]")) depth <- max(0L, depth - 1L)
    else if (ch == delimiter && depth == 0L) {
      parts   <- c(parts, paste(current, collapse = ""))
      current <- character(0)
      next
    }
    current <- c(current, ch)
  }
  parts <- c(parts, paste(current, collapse = ""))
  str_trim(parts)
}

# ---------------------------------------------------------------------------
# extract_leading_label: text label before numbers in a segment
# e.g. "Men: 62/176" -> "Men"
# ---------------------------------------------------------------------------
extract_leading_label <- function(text) {
  m <- str_match(text, "^([A-Za-z][A-Za-z\\s\\-_]*?)\\s*[:/]?\\s*(?=\\d)")
  if (!is.na(m[1,1])) {
    lbl <- str_trim(str_remove(m[1,2], ":$"))
    if (nchar(lbl) > 0 && nchar(lbl) < 40) return(lbl)
  }
  NA_character_
}

# ---------------------------------------------------------------------------
# fix_typo_fractions: fix X/X/Y -> X/Y and X/Y/Y -> X/Y
# ---------------------------------------------------------------------------
fix_typo_fractions <- function(s) {
  # X/X/Y -> X/Y  (first two identical)
  s <- str_replace_all(s,
                       "(\\d+\\.?\\d*)\\s*/\\s*\\1\\s*/\\s*(\\d+\\.?\\d*)", "\\1/\\2")
  # X/Y/Y -> X/Y  (last two identical)
  s <- str_replace_all(s,
                       "(\\d+\\.?\\d*)\\s*/\\s*(\\d+\\.?\\d*)\\s*/\\s*\\2", "\\1/\\2")
  s
}

# ---------------------------------------------------------------------------
# parse_single_block: parse one text block -> list(num, prop, pop, pct, prop_direct)
# ---------------------------------------------------------------------------
parse_single_block <- function(text) {
  out <- list(num = NA_real_, prop = NA_real_, pop = NA_real_,
              pct = NA_real_, prop_direct = NA_real_)
  if (is.na(text) || !str_detect(text, "\\d")) return(out)
  
  s <- fix_typo_fractions(str_trim(text))
  
  # Pattern: N/M (pct%)  e.g. "31/86 (36.0%)" or "n=31/86 (36%)"
  m <- str_match(s, "(\\d+)\\s*/\\s*(\\d+)\\s*[\\(\\s]*(\\d+\\.?\\d*)\\s*%")
  if (!is.na(m[1,1])) {
    out$num  <- as.numeric(m[1,2])
    out$pop  <- as.numeric(m[1,3])
    out$pct  <- as.numeric(m[1,4])
    out$prop <- out$pct / 100
    return(out)
  }
  
  # Pattern: N/M (no percent)  e.g. "31/86"
  m <- str_match(s, "(\\d+)\\s*/\\s*(\\d+)")
  if (!is.na(m[1,1])) {
    out$num <- as.numeric(m[1,2])
    out$pop <- as.numeric(m[1,3])
    return(out)
  }
  
  # Pattern: N (pct%)  e.g. "31 (36.0%)" or "31, 36%"
  m <- str_match(s, "(\\d+)\\s*[,\\(]\\s*(\\d+\\.?\\d*)\\s*%")
  if (!is.na(m[1,1])) {
    out$num  <- as.numeric(m[1,2])
    out$pct  <- as.numeric(m[1,3])
    out$prop <- out$pct / 100
    return(out)
  }
  
  # Pattern: standalone pct%  e.g. "36.0%"
  m <- str_match(s, "(\\d+\\.?\\d*)\\s*%")
  if (!is.na(m[1,1])) {
    out$pct  <- as.numeric(m[1,2])
    out$prop <- out$pct / 100
    return(out)
  }
  
  # Pattern: N=X  e.g. "n=31" or "n = 31"
  m <- str_match(s, regex("n\\s*=\\s*(\\d+)", ignore_case = TRUE))
  if (!is.na(m[1,1])) {
    out$num <- as.numeric(m[1,2])
    return(out)
  }
  
  # All numbers in cell
  all_nums <- as.numeric(str_extract_all(s, "\\d+\\.?\\d*")[[1]])
  
  # Single number
  if (length(all_nums) == 1) {
    v <- all_nums[1]
    if (v >= 0 && v <= 1) {
      # Direct proportion
      out$prop_direct <- v
      out$prop        <- v
    } else if (v == round(v)) {
      out$num <- v
    }
    return(out)
  }
  
  # Two+ numbers: heuristics
  if (length(all_nums) >= 2) {
    v0 <- all_nums[1]; v1 <- all_nums[2]
    
    # First is integer >1, second is decimal 0-1: abs + proportion
    if (v0 > 1 && v0 == round(v0) && v1 >= 0 && v1 <= 1) {
      out$num         <- v0
      out$prop_direct <- v1
      out$prop        <- v1
      return(out)
    }
    
    # First >100 as integer, second <=100: abs + percentage
    if (v0 > 100 && v0 == round(v0) && v1 <= 100) {
      out$num  <- v0
      out$pct  <- v1
      out$prop <- v1 / 100
      return(out)
    }
    
    # First is integer, second > first (unlikely abs/pop without /): standalone abs
    if (v0 == round(v0) && v0 > 0) {
      out$num <- v0
      return(out)
    }
  }
  
  # Fallback: if single integer-like number present, take it as abs
  m2 <- str_match(s, "^[^\\d]*(\\d+)")
  if (!is.na(m2[1,1])) {
    v <- as.numeric(m2[1,2])
    if (v == round(v)) out$num <- v
  }
  
  out
}

# ---------------------------------------------------------------------------
# split_into_blocks: split cell into labelled blocks (one per group or overall)
# Returns list of list(label, text)
# ---------------------------------------------------------------------------
split_into_blocks <- function(text) {
  if (is.na(text) || !str_detect(text, "\\d"))
    return(list(list(label = "overall", text = text %||% "")))
  
  s <- str_replace_all(str_trim(text), "\\s+", " ")
  
  # Try semicolon split outside parens
  semi_parts <- split_outside_parens(s, ";")
  semi_parts <- semi_parts[nchar(semi_parts) > 0]
  if (length(semi_parts) > 1 && all(sapply(semi_parts, function(p) str_detect(p, "\\d")))) {
    return(lapply(seq_along(semi_parts), function(i) {
      lbl <- extract_leading_label(semi_parts[i])
      list(label = if (!is.na(lbl)) lbl else paste0("group_", i),
           text  = semi_parts[i])
    }))
  }
  
  # Try "vs" split
  vs_parts <- str_trim(str_split(s, regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
  vs_parts <- vs_parts[nchar(vs_parts) > 0]
  if (length(vs_parts) == 2 && all(sapply(vs_parts, function(p) str_detect(p, "\\d")))) {
    return(lapply(seq_along(vs_parts), function(i) {
      lbl <- extract_leading_label(vs_parts[i])
      list(label = if (!is.na(lbl)) lbl else paste0("group_", i),
           text  = vs_parts[i])
    }))
  }
  
  list(list(label = "overall", text = s))
}

`%||%` <- function(a, b) if (!is.null(a) && !is.na(a)) a else b

# ---------------------------------------------------------------------------
# parse_proportion_cell: master parsing function for a single cell
# Returns a named list of all extracted fields and flags
# ---------------------------------------------------------------------------
AVG_KEYWORDS_PAT <- regex(
  "\\bmedian\\b|\\bmean\\b|\\bsd\\b|\\biqr\\b|\\blqr\\b|\\buqr\\b|\\+/-|±",
  ignore_case = TRUE)

parse_proportion_cell <- function(cell_val) {
  empty <- list(
    abs_num_reported  = "N", abs_frac_reported = "N",
    pct_reported      = "N", prop_reported     = "N",
    overall_num = NA_real_, overall_prop = NA_real_, overall_pop = NA_real_,
    sub1_label  = NA_character_, sub1_num = NA_real_,
    sub1_prop   = NA_real_,      sub1_pop = NA_real_,
    sub2_label  = NA_character_, sub2_num = NA_real_,
    sub2_prop   = NA_real_,      sub2_pop = NA_real_,
    flag_zero_one    = "N", flag_not          = "N",
    flag_question    = "N", flag_avg_keyword  = "N",
    raw_nums         = numeric(0),
    parsed_blocks    = list()
  )
  
  if (is.na(cell_val) || str_trim(as.character(cell_val)) == "")
    return(empty)
  
  s <- normalise_text(as.character(cell_val))
  if (is.na(s)) return(empty)
  
  empty$flag_not         <- if (str_detect(s, regex("\\bnot\\b", ignore_case=TRUE))) "Y" else "N"
  empty$flag_question    <- if (str_detect(s, "\\?")) "Y" else "N"
  empty$flag_avg_keyword <- if (str_detect(s, AVG_KEYWORDS_PAT)) "Y" else "N"
  empty$raw_nums         <- as.numeric(str_extract_all(s, "\\d+\\.?\\d*")[[1]])
  
  blocks <- split_into_blocks(s)
  parsed <- lapply(blocks, function(b) {
    p        <- parse_single_block(b$text)
    p$label  <- b$label
    p
  })
  empty$parsed_blocks <- parsed
  
  # Assign overall vs subgroup blocks
  if (length(parsed) == 0) return(empty)
  
  if (length(parsed) == 1 || parsed[[1]]$label == "overall") {
    ov <- parsed[[1]]
    subs <- if (length(parsed) > 1) parsed[-1] else list()
  } else {
    ov   <- NULL
    subs <- parsed
  }
  
  # Overall values
  if (!is.null(ov)) {
    empty$overall_num  <- ov$num
    empty$overall_prop <- ov$prop
    empty$overall_pop  <- ov$pop
    if (!is.na(ov$num))         empty$abs_num_reported  <- "Y"
    if (!is.na(ov$pop))         empty$abs_frac_reported <- "Y"
    if (!is.na(ov$pct))         empty$pct_reported      <- "Y"
    if (!is.na(ov$prop_direct)) empty$prop_reported     <- "Y"
  }
  
  # Subgroup 1
  if (length(subs) >= 1) {
    empty$sub1_label <- subs[[1]]$label
    empty$sub1_num   <- subs[[1]]$num
    empty$sub1_prop  <- subs[[1]]$prop
    empty$sub1_pop   <- subs[[1]]$pop
  }
  # Subgroup 2
  if (length(subs) >= 2) {
    empty$sub2_label <- subs[[2]]$label
    empty$sub2_num   <- subs[[2]]$num
    empty$sub2_prop  <- subs[[2]]$prop
    empty$sub2_pop   <- subs[[2]]$pop
  }
  
  # Zero/one flag
  pv <- empty$overall_prop %||% empty$sub1_prop
  if (!is.null(pv) && !is.na(pv) && (pv == 0 || pv == 1))
    empty$flag_zero_one <- "Y"
  
  empty
}

# ---------------------------------------------------------------------------
# safe_eq: approximate equality check (tolerance 1%)
# ---------------------------------------------------------------------------
safe_eq <- function(a, b, tol = 0.01) {
  if (is.null(a) || is.null(b) || is.na(a) || is.na(b)) return(NA_character_)
  fa <- suppressWarnings(as.numeric(a))
  fb <- suppressWarnings(as.numeric(b))
  if (is.na(fa) || is.na(fb)) return(NA_character_)
  if (fb == 0) return(if (fa == 0) "Y" else "N")
  if (abs(fa - fb) / max(abs(fb), 1e-9) <= tol) "Y" else "N"
}

# ---------------------------------------------------------------------------
# flag_duplication: Y if numbers from this cell appear in other prop cols
# ---------------------------------------------------------------------------
flag_duplication <- function(cell_nums, row_idx, current_col_names, df,
                             all_prop_col_names) {
  if (length(cell_nums) == 0) return("N")
  other_cols <- setdiff(all_prop_col_names, current_col_names)
  other_cols <- intersect(other_cols, colnames(df))
  other_nums <- numeric(0)
  for (col in other_cols) {
    val <- df[[col]][row_idx]
    if (!is.na(val)) {
      found <- as.numeric(str_extract_all(as.character(val), "\\d+\\.?\\d*")[[1]])
      other_nums <- c(other_nums, found)
    }
  }
  overlap <- intersect(round(cell_nums, 4), round(other_nums, 4))
  if (length(overlap) > 0) "Y" else "N"
}

# ---------------------------------------------------------------------------
# build_study_reference: "Surname Year"
# ---------------------------------------------------------------------------
build_study_reference <- function(author_raw, year_raw) {
  # FIX: use isTRUE(!is.na(...)) rather than bare !is.na() to safely handle
  # NULL or zero-length values that would otherwise cause
  # "argument is of length zero" when passed to if()
  author <- if (isTRUE(!is.na(author_raw)) && length(author_raw) > 0)
    str_trim(as.character(author_raw)) else ""
  year   <- if (isTRUE(!is.na(year_raw))   && length(year_raw)   > 0)
    str_trim(as.character(year_raw))   else ""
  # First author: split on comma or 'and', take first; then last word = surname
  first_author <- str_trim(str_split(author, ",|\\band\\b")[[1]][1])
  surname      <- str_trim(tail(str_split(first_author, "\\s+")[[1]], 1))
  str_trim(paste(surname, year))
}

# =============================================================================
# PART 1: Clean column S (number of participants)
# =============================================================================
df <- read.xlsx(input_path, sheet = 1)

n_total_vec <- sapply(df[[COL_N_TOTAL]], clean_n_total)

# =============================================================================
# PART 2: Build full set of proportion column names for duplication check
# =============================================================================
all_prop_col_names <- unique(unlist(lapply(PROPORTION_COLS, function(x) x[[2]])))
all_prop_col_names <- intersect(all_prop_col_names, colnames(df))

# =============================================================================
# PART 3: Process each proportion column
# =============================================================================
all_sheet_data <- list()

for (col_def in PROPORTION_COLS) {
  sheet_name    <- col_def[[1]]
  src_col_names <- col_def[[2]]
  cat("Processing", sheet_name, "...\n")
  
  # Build combined series: first non-NA across source columns
  existing_cols <- intersect(src_col_names, colnames(df))
  if (length(existing_cols) == 0) {
    message("  No source columns found for ", sheet_name, " -- skipping.")
    next
  }
  
  combined_vals <- apply(df[, existing_cols, drop = FALSE], 1, function(row) {
    v <- row[!is.na(row) & row != ""]
    if (length(v) == 0) NA_character_ else as.character(v[1])
  })
  
  rows_out <- vector("list", nrow(df))
  
  for (i in seq_len(nrow(df))) {
    row       <- df[i, ]
    cell_val  <- combined_vals[i]
    n_total   <- n_total_vec[i]
    
    study_num <- row[[COL_STUDY_NUM]]
    study_ref <- build_study_reference(row[[COL_AUTHORS]], row[[COL_YEAR]])
    
    # Subgroup detection
    sub_info <- detect_subgroups(
      row[[COL_GROUPS]], row[[COL_GROUP1]], row[[COL_GROUP2]])
    
    # Parse cell
    parsed <- parse_proportion_cell(cell_val)
    
    # Convenience aliases
    ov_num  <- parsed$overall_num
    ov_prop <- parsed$overall_prop
    ov_pop  <- parsed$overall_pop
    
    # Primary value = overall if present, else subgroup 1
    prim_num  <- if (!is.na(ov_num))  ov_num  else parsed$sub1_num
    prim_prop <- if (!is.na(ov_prop)) ov_prop else parsed$sub1_prop
    prim_pop  <- if (!is.na(ov_pop))  ov_pop  else parsed$sub1_pop
    
    # Validation checks
    manual_prop <- if (!is.na(prim_num) && !is.na(n_total) && n_total > 0)
      prim_num / n_total else NA_real_
    prop_equal  <- safe_eq(manual_prop, prim_prop)
    
    manual_num  <- if (!is.na(prim_prop) && !is.na(n_total))
      prim_prop * n_total else NA_real_
    num_equal   <- safe_eq(manual_num, prim_num)
    
    pop_equal   <- safe_eq(prim_pop, n_total)
    
    # Unused numbers flag
    used_nums <- numeric(0)
    for (b in parsed$parsed_blocks) {
      for (field in c("num", "pop")) {
        v <- b[[field]]
        if (!is.null(v) && !is.na(v)) used_nums <- c(used_nums, v)
      }
      if (!is.null(b$pct) && !is.na(b$pct)) {
        used_nums <- c(used_nums, b$pct, b$pct / 100)
      }
      if (!is.null(b$prop_direct) && !is.na(b$prop_direct))
        used_nums <- c(used_nums, b$prop_direct)
    }
    raw_nums    <- parsed$raw_nums
    unused      <- raw_nums[!sapply(raw_nums, function(n)
      any(abs(n - used_nums) < 0.001))]
    flag_unused <- if (length(unused) > 0) "Y" else "N"
    
    # Duplication flag
    flag_dup <- flag_duplication(
      raw_nums, i, existing_cols, df, all_prop_col_names)
    
    # Assemble output row in specified column order:
    # study number > study reference > flags > reported/values > checks
    rows_out[[i]] <- list(
      study_number        = study_num,
      study_reference     = study_ref,
      # --- FLAGS ---
      flag_not            = parsed$flag_not,
      flag_question       = parsed$flag_question,
      flag_avg_keyword    = parsed$flag_avg_keyword,
      flag_unused_numbers = flag_unused,
      flag_zero_one       = parsed$flag_zero_one,
      flag_duplication    = flag_dup,
      has_subgroups       = if (sub_info$has_subgroups) "Y" else "N",
      n_subgroups         = sub_info$n_subgroups,
      # --- REPORTED TYPES ---
      abs_num_reported    = parsed$abs_num_reported,
      abs_frac_reported   = parsed$abs_frac_reported,
      pct_reported        = parsed$pct_reported,
      prop_reported       = parsed$prop_reported,
      # --- OVERALL VALUES ---
      overall_num         = ov_num,
      overall_prop        = if (!is.na(ov_prop)) round(ov_prop, 6) else NA_real_,
      overall_pop         = ov_pop,
      # --- SUBGROUP 1 ---
      subgroup1_label     = parsed$sub1_label,
      subgroup1_num       = parsed$sub1_num,
      subgroup1_prop      = if (!is.na(parsed$sub1_prop))
        round(parsed$sub1_prop, 6) else NA_real_,
      subgroup1_pop       = parsed$sub1_pop,
      # --- SUBGROUP 2 ---
      subgroup2_label     = parsed$sub2_label,
      subgroup2_num       = parsed$sub2_num,
      subgroup2_prop      = if (!is.na(parsed$sub2_prop))
        round(parsed$sub2_prop, 6) else NA_real_,
      subgroup2_pop       = parsed$sub2_pop,
      # --- CHECKS ---
      manual_prop         = if (!is.na(manual_prop)) round(manual_prop, 6) else NA_real_,
      prop_equal          = prop_equal,
      manual_num          = if (!is.na(manual_num))  round(manual_num,  2)  else NA_real_,
      num_equal           = num_equal,
      pop_equal           = pop_equal
    )
  }
  
  all_sheet_data[[sheet_name]] <- bind_rows(rows_out)
}

# =============================================================================
# PART 4: Write output workbooks
# =============================================================================

FLAG_COLS  <- c("flag_not","flag_question","flag_avg_keyword","flag_unused_numbers",
                "flag_zero_one","flag_duplication","has_subgroups","n_subgroups")
CHECK_COLS <- c("manual_prop","prop_equal","manual_num","num_equal","pop_equal")

write_workbook <- function(sheet_data_list, path, fill_nr = FALSE) {
  wb <- createWorkbook()
  
  header_style <- createStyle(fontName = "Arial", fontSize = 10,
                              textDecoration = "bold",
                              fgFill = "#D9E1F2", wrapText = TRUE,
                              border = "Bottom")
  flag_style   <- createStyle(fgFill = "#FFF2CC")
  check_style  <- createStyle(fgFill = "#E2EFDA")
  
  for (sheet_name in names(sheet_data_list)) {
    df_sheet <- sheet_data_list[[sheet_name]]
    
    if (fill_nr) {
      df_sheet <- df_sheet %>%
        mutate(across(everything(), ~ {
          x <- .x
          x[is.na(x) | x == ""] <- "NR"
          x
        }))
    }
    
    addWorksheet(wb, sheetName = substr(sheet_name, 1, 31))
    writeData(wb, sheet = sheet_name, x = df_sheet, headerStyle = header_style)
    
    n_rows <- nrow(df_sheet)
    if (n_rows > 0) {
      for (j in seq_along(colnames(df_sheet))) {
        col_h <- colnames(df_sheet)[j]
        if (col_h %in% FLAG_COLS) {
          addStyle(wb, sheet = sheet_name,
                   style = flag_style,
                   rows = 2:(n_rows + 1), cols = j, gridExpand = TRUE)
        } else if (col_h %in% CHECK_COLS) {
          addStyle(wb, sheet = sheet_name,
                   style = check_style,
                   rows = 2:(n_rows + 1), cols = j, gridExpand = TRUE)
        }
      }
    }
    
    # Approximate column widths
    for (j in seq_along(colnames(df_sheet))) {
      setColWidths(wb, sheet = sheet_name, cols = j,
                   widths = max(12, nchar(colnames(df_sheet)[j]) + 2))
    }
  }
  
  saveWorkbook(wb, path, overwrite = TRUE)
  cat("Saved:", path, "\n")
}

write_workbook(all_sheet_data, output_path,    fill_nr = FALSE)
write_workbook(all_sheet_data, output_path_nr, fill_nr = TRUE)

cat("Done.\n")