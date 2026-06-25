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
# NEW IN THIS VERSION:
#   - any_flag_string: names the specific flag/check column(s) that
#     triggered for this row (e.g. "flag_zero_one, subgroup1_pop_equal")
#   - flag_excess_subgroups: "Y" if more than 2 subgroups were identified
#     (this script only has columns for up to 2 named subgroups)
#   - all_subgroup_names: combines subgroup names from columns BB, BC, T,
#     and any names found directly in the individual proportion cell
#   - A third output workbook (proportion_cleaned_condensed.xlsx) with only
#     a condensed subset of columns -- see CONDENSED_COLS below
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
output_path     <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v6_add.xlsx"
output_path_nr  <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v6_add_NR.xlsx"

# NEW FEATURE: third, separate workbook containing only a condensed subset
# of columns (see CONDENSED_COLS below for the exact list)
output_path_condensed <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v6_add_condensed.xlsx"

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
COL_AGE        <- "age"                     # R -- FIX (study 101, 111):
# used by flag_duplication so a
# "number of participants" value
# accidentally copy-pasted from
# the age column is caught

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
# is_short_label_list: helper used only by detect_subgroups (below) to decide
# whether commas in column T are acting as group-name separators (e.g.
# "survivors, non-survivors", "COPD, non COPD") rather than ordinary
# punctuation within a longer narrative sentence (e.g. "Among patients in
# the study cohort, 15% (32 of 217) died; of ward-triage patients, ...").
#
# A blanket rule treating every comma as a separator was tested and found to
# badly over-split many narrative cells elsewhere in column T (e.g. turning
# a sensible 2-part split into a meaningless 6-7 part split). This scoped
# heuristic only accepts the comma-split when the WHOLE string is short,
# splits into 2-4 parts, and each part reads like a short label (a handful
# of words, no sentence punctuation) -- which safely matches study 127, 131,
# 132 type cells without affecting any longer descriptive cell.
# ---------------------------------------------------------------------------
is_short_label_list <- function(g) {
  if (nchar(g) > 60) return(FALSE)
  parts <- str_trim(str_split(g, ",")[[1]])
  parts <- parts[nchar(parts) > 0]
  if (length(parts) < 2 || length(parts) > 4) return(FALSE)
  all(sapply(parts, function(p) {
    n_words <- length(str_split(p, "\\s+")[[1]])
    n_words <= 5 && !str_detect(p, "[.;]")
  }))
}

# ---------------------------------------------------------------------------
# detect_subgroups: uses columns T, BB, BC to detect subgroups
# Returns list(has_subgroups=TRUE/FALSE, n_subgroups=integer or NA,
#              names_from_T=character vector of names parsed from column T,
#              name_from_BB=character or NA, name_from_BC=character or NA)
#
# NEW FEATURE: names_from_T/name_from_BB/name_from_BC added so the calling
# code can build the "all_subgroup_names" column, combining names found in
# the study-design columns (T, BB, BC) with any names found directly in the
# individual proportion cell being processed (sub1_label/sub2_label).
#
# FIX (study 127, 131, 132): column T sometimes uses a plain comma to
# separate a short list of group names with no other contrast keyword
# present, e.g. "survivors, non-survivors", "COPD, non COPD", "admissions
# 1995-1999, 1999-2004". These are now detected via is_short_label_list()
# and split on comma -- but ONLY when that narrow check passes, so commas
# inside longer narrative text elsewhere in column T are left untouched.
# ---------------------------------------------------------------------------
CONTRAST_PAT <- regex(
  "\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|\\bv\\b|\\band\\b|;|/|\\bgroup\\b",
  ignore_case = TRUE)

detect_subgroups <- function(groups_val, g1_val, g2_val) {
  g  <- if (isTRUE(!is.na(groups_val)) && length(groups_val) > 0) str_trim(as.character(groups_val)) else ""
  g1 <- if (isTRUE(!is.na(g1_val))    && length(g1_val)    > 0) str_trim(as.character(g1_val))     else ""
  g2 <- if (isTRUE(!is.na(g2_val))    && length(g2_val)    > 0) str_trim(as.character(g2_val))     else ""
  
  # FIX: check the narrow short-label-list comma case FIRST, before the
  # "single" exclusion below, so e.g. "COPD, non COPD" is correctly
  # detected even though it contains no other contrast keyword.
  if (nchar(g) > 0 && is_short_label_list(g)) {
    parts <- str_trim(str_split(g, ",")[[1]])
    parts <- parts[nchar(parts) > 0]
    return(list(has_subgroups = TRUE, n_subgroups = length(parts),
                names_from_T = parts, name_from_BB = NA_character_,
                name_from_BC = NA_character_))
  }
  
  # "single" without contrast => no subgroups
  if (str_detect(g, regex("\\bsingle\\b", ignore_case = TRUE)) &&
      !str_detect(g, CONTRAST_PAT)) {
    return(list(has_subgroups = FALSE, n_subgroups = 0L,
                names_from_T = character(0), name_from_BB = NA_character_,
                name_from_BC = NA_character_))
  }
  
  # BB and BC both populated => 2 subgroups (named directly by BB/BC)
  if (nchar(g1) > 0 && nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L,
                names_from_T = character(0), name_from_BB = g1, name_from_BC = g2))
  if (nchar(g1) > 0 || nchar(g2) > 0)
    return(list(has_subgroups = TRUE, n_subgroups = 2L,
                names_from_T = character(0),
                name_from_BB = if (nchar(g1) > 0) g1 else NA_character_,
                name_from_BC = if (nchar(g2) > 0) g2 else NA_character_))
  
  # Count parts split by vs/;
  parts <- str_trim(str_split(g, regex("\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|;",
                                       ignore_case = TRUE))[[1]])
  parts <- parts[nchar(parts) > 0]
  if (length(parts) >= 2)
    return(list(has_subgroups = TRUE, n_subgroups = length(parts),
                names_from_T = parts, name_from_BB = NA_character_,
                name_from_BC = NA_character_))
  
  if (str_detect(g, CONTRAST_PAT))
    return(list(has_subgroups = TRUE, n_subgroups = NA_integer_,
                names_from_T = character(0), name_from_BB = NA_character_,
                name_from_BC = NA_character_))
  
  list(has_subgroups = FALSE, n_subgroups = 0L,
       names_from_T = character(0), name_from_BB = NA_character_,
       name_from_BC = NA_character_)
}

# ---------------------------------------------------------------------------
# combine_subgroup_names: NEW FEATURE -- builds the "all_subgroup_names"
# column by merging subgroup names from every available source:
#   - column BB (intervention group) and BC (control group) -- the study's
#     named arms, if directly stated
#   - column T (groups if applicable) -- names parsed out of the free-text
#     description of the study's groups
#   - the individual proportion cell itself -- subgroup1_label/
#     subgroup2_label, which may name a subgroup the study-design columns
#     don't (e.g. "Men"/"Women" appearing only in the icu mortality cell)
# Names are deduplicated (case-insensitively) and joined with "; " so the
# same name appearing in multiple sources (e.g. both column T and the cell
# itself) is only listed once.
# ---------------------------------------------------------------------------
combine_subgroup_names <- function(names_from_T, name_from_BB, name_from_BC,
                                   cell_sub1_label, cell_sub2_label) {
  all_names <- c(names_from_T, name_from_BB, name_from_BC,
                 cell_sub1_label, cell_sub2_label)
  all_names <- all_names[!is.na(all_names)]
  all_names <- str_trim(all_names)
  all_names <- all_names[nchar(all_names) > 0]
  if (length(all_names) == 0) return(NA_character_)
  
  # Cosmetic cleanup: if a generic "group_N" placeholder (from a source
  # with no real name available, e.g. study 173's unlabelled "50%+16%"
  # split) appears ALONGSIDE a genuinely named alternative from a different
  # source for the same study (e.g. column T's "development"/"validation"
  # for study 4, while the cell itself only produced "group_1"/"group_2"),
  # the generic placeholder is redundant and is dropped. If NO named
  # alternative exists anywhere, the generic placeholder is kept, since it
  # is the only information available.
  is_generic <- str_detect(all_names, regex("^group_\\d+$", ignore_case = TRUE))
  if (any(is_generic) && any(!is_generic)) {
    all_names <- all_names[!is_generic]
  }
  
  # Deduplicate case-insensitively while preserving first-seen casing
  keep <- !duplicated(tolower(all_names))
  all_names <- all_names[keep]
  paste(all_names, collapse = "; ")
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
# extract_leading_label: text label for a segment, checking LEADING position
# first (e.g. "Men: 62/176" -> "Men"), then falling back to TRAILING
# position if no leading label is found (e.g. "81/185 (44%) overall" ->
# "overall", "72/156 (46%) indigenous" -> "indigenous").
#
# FIX (study 199): "81/185 (44%) overall; 72/156 (46%) indigenous; 9/29
# (31%) non-indigenous" -- all three segments here have their label AFTER
# the numbers, not before. The function previously only checked the
# leading position, so all three fell through to a generic "group_N" name,
# and "overall" specifically was never recognised as the literal overall
# block (it was just treated as another named subgroup, "1 subgroup
# missing" from the user's report, since "overall" effectively displaced a
# genuine subgroup slot).
# ---------------------------------------------------------------------------
extract_leading_label <- function(text) {
  m <- str_match(text, "^([A-Za-z][A-Za-z\\s\\-_]*?)\\s*[:/]?\\s*(?=\\d)")
  if (!is.na(m[1,1])) {
    lbl <- str_trim(str_remove(m[1,2], ":$"))
    if (nchar(lbl) > 0 && nchar(lbl) < 40) return(lbl)
  }
  # FIX (study 198 regression guard): the trailing-label fallback below is
  # unsafe when a single segment secretly contains TWO groups' worth of
  # data joined by "and" (e.g. "11 (21%) in the treatment group and 17
  # (37%) in the control group" -- the segment as a whole was only parsed
  # as ONE block with num=11/pct=21%, the FIRST group's figures, but the
  # trailing label would incorrectly grab "in the control group" from the
  # SECOND group's text at the end). Detected by looking for a percentage
  # ... "and" ... another percentage pattern within the same segment; if
  # found, the trailing-label fallback is skipped entirely so the caller's
  # safe "group_N" fallback is used instead, rather than attaching a
  # plausible-looking but WRONG label to the wrong figures.
  has_and_joined_pair <- str_detect(text, regex(
    "\\d+\\.?\\d*\\s*%.*?\\band\\b.*?\\d+\\.?\\d*\\s*%", ignore_case = TRUE))
  if (has_and_joined_pair) return(NA_character_)
  
  # Trailing label fallback: text ends in a short word/phrase with no
  # further digits after it (so it reads as a label describing the figures
  # that came before it, not a continuation of the data).
  m2 <- str_match(text, "(?:^|[\\)\\]%])\\s*([A-Za-z][A-Za-z\\s\\-]*)\\s*$")
  if (!is.na(m2[1,1])) {
    lbl <- str_trim(m2[1,2])
    if (nchar(lbl) > 0 && nchar(lbl) < 40) return(lbl)
  }
  NA_character_
}

# ---------------------------------------------------------------------------
# convert_word_numbers: convert spelled-out cardinal numbers (one, two,
# nine, twenty, etc.) to digits.
#
# FIX (study 95): "Nine of these patients required mechanic ventilation."
# -- the count is spelled out as a word rather than written as a digit, so
# none of the numeric-pattern matching below could find it at all. This
# converts common cardinal number words (0-20, plus the tens up to ninety)
# to their digit form before any other parsing happens. Scoped to whole
# cardinal numbers only (not ordinals like "ninth", and not fractions of
# words like "nineteen" colliding with "nine" -- word boundaries and
# matching the longer words first prevent this).
# ---------------------------------------------------------------------------
NUMBER_WORDS <- c(
  "nineteen"=19,"eighteen"=18,"seventeen"=17,"sixteen"=16,"fifteen"=15,
  "fourteen"=14,"thirteen"=13,"eleven"=11,"twelve"=12,"twenty"=20,
  "thirty"=30,"forty"=40,"fifty"=50,"sixty"=60,"seventy"=70,"eighty"=80,
  "ninety"=90,"zero"=0,"one"=1,"two"=2,"three"=3,"four"=4,"five"=5,
  "six"=6,"seven"=7,"eight"=8,"nine"=9,"ten"=10
)

convert_word_numbers <- function(text) {
  for (word in names(NUMBER_WORDS)) {
    text <- str_replace_all(text, regex(paste0("\\b", word, "\\b"), ignore_case = TRUE),
                            as.character(NUMBER_WORDS[[word]]))
  }
  text
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
  if (is.na(text)) return(out)
  
  # FIX (study 6): "All patients" (or just "All") means the proportion is
  # 100% (prop=1), even though the cell contains no digit at all. Checked
  # before the digit-existence early return below, since this case has no
  # digit. Scoped to the text STARTING with "all" (optionally followed by a
  # describing word like "patients") and containing nothing else, so it
  # does not misfire on unrelated uses of "all" elsewhere in a sentence.
  # A symmetric "none"/"no patients" case is also handled, giving prop=0.
  text_trimmed <- str_trim(text)
  if (str_detect(text_trimmed, regex("^all(\\s+[a-zA-Z]+)?\\.?$", ignore_case = TRUE))) {
    out$prop_direct <- 1
    out$prop        <- 1
    return(out)
  }
  if (str_detect(text_trimmed, regex("^(none|no patients|no one|nobody)(\\s+[a-zA-Z]+)*\\.?$", ignore_case = TRUE))) {
    out$prop_direct <- 0
    out$prop        <- 0
    return(out)
  }
  
  # FIX (study 95): convert spelled-out numbers to digits BEFORE checking
  # whether the text contains any digit at all -- "Nine of these patients"
  # has no digit until "Nine" is converted to "9".
  text <- convert_word_numbers(text)
  if (!str_detect(text, "\\d")) return(out)
  
  s <- fix_typo_fractions(str_trim(text))
  
  # FIX (study 131, 132): collapse a doubled percent sign typo "%%" -> "%"
  # e.g. "49%%" -> "49%", "25.2%%" -> "25.2%". Done before any pattern
  # matching so every percentage pattern below benefits automatically.
  s <- str_replace_all(s, "%%+", "%")
  
  # FIX (#84): strip thousands-separator commas from inside numbers
  # e.g. "1,166" -> "1166" so the N/M and percentage patterns can match
  # across what would otherwise be a broken digit sequence.
  # Only strips a comma that sits BETWEEN digits (comma-digit pattern),
  # so commas used as general punctuation/list separators are untouched.
  s <- str_replace_all(s, "(?<=\\d),(?=\\d{3}(\\D|$))", "")
  
  # FIX (study 1): strip a trailing "n, (%)" / "n (%)" / "both n, (%)" style
  # label. This dataset repeatedly uses this as a HEADER-STYLE notation
  # meaning "the preceding figures are an n (%) pair" rather than the "%"
  # symbol being attached to a specific number (e.g. "49, (57) n, (%)" means
  # n=49, 57%, with "n, (%)" just describing the format used). Verified
  # against every column in the dataset: this trailing label appears after
  # 7 different cells, all of which follow this same header-style
  # convention, so stripping it before the normal pattern cascade below is
  # safe and recovers the percentage that would otherwise be missed.
  s <- str_remove(s, regex("\\s*(?:both\\s+)?n\\s*,?\\s*\\(\\s*%\\s*\\)\\s*$", ignore_case = TRUE))
  s <- str_trim(s)
  
  # FIX (study 201, 204): percentage appears BEFORE the fraction, e.g.
  # "27.9% (1058/3786)" -- previously only the N/M...pct% ORDER was matched,
  # so this format fell through with num/pop captured by a later pattern but
  # pct/prop never set at all. This pattern is tried first since it is more
  # specific (anchored on % immediately followed by a bracketed fraction).
  m <- str_match(s, "(\\d+\\.?\\d*)\\s*%\\s*\\(\\s*(\\d+)\\s*/\\s*(\\d+)\\s*\\)")
  if (!is.na(m[1, 1])) {
    out$pct  <- as.numeric(m[1, 2])
    out$prop <- out$pct / 100
    out$num  <- as.numeric(m[1, 3])
    out$pop  <- as.numeric(m[1, 4])
    return(out)
  }
  
  # Pattern: N/M (pct%)  e.g. "31/86 (36.0%)" or "n=31/86 (36%)"
  # Also now matches when % is separated by extra spaces e.g. "722/1166  61.9%"
  # FIX (#84): "[\\(\\s]*" widened so a percentage following the fraction with
  # no bracket at all (just whitespace) is still captured here.
  m <- str_match(s, "(\\d+)\\s*/\\s*(\\d+)\\s*[\\(\\s]*(\\d+\\.?\\d*)\\s*%")
  if (!is.na(m[1,1])) {
    out$num  <- as.numeric(m[1,2])
    out$pop  <- as.numeric(m[1,3])
    out$pct  <- as.numeric(m[1,4])
    out$prop <- out$pct / 100
    return(out)
  }
  
  # FIX (#81,#82): N/M (decimal) with NO % sign inside the brackets
  # e.g. "14/77 (18.2)" -- the bracketed decimal is the percentage, the %
  # symbol was simply omitted by the original study/extractor.
  # Guarded so this only fires when the bracketed number is plausible as a
  # percentage point value (0-100) and is NOT also matched by the pop-bound
  # case below (handled by pattern order: this runs only after the explicit
  # "%"-bearing pattern above has already failed to match).
  m <- str_match(s, "(\\d+)\\s*/\\s*(\\d+)\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
  if (!is.na(m[1,1])) {
    n   <- as.numeric(m[1,2]); pop <- as.numeric(m[1,3]); bracketed <- as.numeric(m[1,4])
    out$num <- n
    out$pop <- pop
    if (bracketed >= 0 && bracketed <= 100) {
      out$pct  <- bracketed
      out$prop <- bracketed / 100
    }
    return(out)
  }
  
  # Pattern: N/M (no percent, no bracket at all)  e.g. "31/86"
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
  
  # FIX (#81,#82 single-number variant): N (decimal) with no % sign
  # e.g. "278 (23.8)" -- same omitted-% issue as above but without a fraction
  # FIX (study 1): comma allowed between the number and bracket, e.g.
  # "49, (57)" (after the trailing "n, (%)" label has been stripped above)
  m <- str_match(s, "(\\d+)\\s*,?\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)")
  if (!is.na(m[1,1])) {
    n <- as.numeric(m[1,2]); bracketed <- as.numeric(m[1,3])
    if (bracketed >= 0 && bracketed <= 100) {
      out$num  <- n
      out$pct  <- bracketed
      out$prop <- bracketed / 100
      return(out)
    }
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
      # Whole number -> absolute count
      out$num <- v
    } else if (v > 1 && v <= 100) {
      # FIX (study 195, 196): isolated decimal number greater than 1 (and
      # not a whole number) is treated as a percentage point value with no
      # "%" symbol present, per the correction rule: "if isolated number +
      # decimal + >1, likely percentage". e.g. "10.6" -> 10.6%, "24.76" -> 24.76%.
      # Bounded at <=100 so an isolated decimal outside the percentage range
      # (e.g. an age or a count with a typo'd decimal) is not misclassified.
      out$pct  <- v
      out$prop <- v / 100
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
    
    # First >100 as integer, second a genuine percentage-looking value:
    # FIX (study 166): "182 total (91 SOC, 91 CPA)" was being misread as
    # num=182, pct=91 (-> prop=0.91) because the old check only required
    # v1 <= 100, and 91 satisfies that even though it is clearly a COUNT of
    # patients here, not a percentage. A bare integer second value with no
    # decimal point and no nearby "%" is far more likely to be another count
    # (as in this cell) than a percentage, so the check now requires v1 to
    # either have a decimal point (e.g. 61.9) or be small enough that it is
    # implausible as a count occurring straight after a >100 total in this
    # specific two-number-only context (<=20, a conservative cutoff chosen
    # so genuine percentages like "61.9" or "5" still match while ordinary
    # subgroup counts like "91" do not).
    v1_has_decimal <- v1 != round(v1)
    if (v0 > 100 && v0 == round(v0) && (v1_has_decimal || v1 <= 20)) {
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
  if (is.na(text)) return(list(list(label = "overall", text = text %||% "")))
  
  # FIX (study 95): convert spelled-out numbers to digits before the
  # digit-presence check below, for the same reason as in
  # parse_single_block -- "Nine of these patients..." has no digit at all
  # until "Nine" is converted to "9".
  text <- convert_word_numbers(text)
  if (!str_detect(text, "\\d"))
    return(list(list(label = "overall", text = text %||% "")))
  
  s <- str_replace_all(str_trim(text), "\\s+", " ")
  
  # FIX (study 4, and study 7 -- generalised further below): "total n=125/575
  # (development cohort: 97/455 (21.3%); external validation cohort: 28/120
  # (23.3%))" -- the OVERALL fraction (125/575) is immediately followed by a
  # bracket that belongs to a nested subgroup breakdown, not its own
  # percentage. Without this check, the general N/M-then-% pattern used
  # later skips past 125/575 entirely and matches the first subgroup
  # fraction (97/455) instead, so the overall value was being lost and a
  # subgroup value wrongly treated as overall.
  #
  # FIX (study 7): "n=101 / 257 (39.3%) (Men: 62 / 176 (35%); Women: 39 / 81
  # (48%))" has the SAME nested-breakdown shape, but differs in two ways:
  # (a) it uses a bare "n=" prefix instead of "total"/"overall", and (b) the
  # overall fraction has its OWN percentage bracket "(39.3%)" before the
  # nested breakdown bracket. The keyword alternation was widened to accept
  # "n" as well, and an optional "(pct%)" bracket is now allowed between the
  # fraction and the nested breakdown bracket. This was verified against
  # every column in the dataset to confirm it ONLY matches these two cells
  # and no other "N/M (...)" cell anywhere else.
  overall_lead <- str_match(s, regex(
    "^(?:total|overall|n)\\s*n?\\s*[=:]?\\s*(\\d+)\\s*/\\s*(\\d+)\\s*(?:\\([\\d.]+%\\)\\s*)?\\((.*;.*)\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(overall_lead[1, 1])) {
    overall_text <- paste0(overall_lead[1, 2], "/", overall_lead[1, 3])
    inner_text   <- overall_lead[1, 4]
    inner_parts  <- split_outside_parens(inner_text, ";")
    inner_parts  <- inner_parts[nchar(str_trim(inner_parts)) > 0]
    
    if (length(inner_parts) > 1 &&
        all(sapply(inner_parts, function(p) str_detect(p, "\\d")))) {
      blocks <- list(list(label = "overall", text = overall_text))
      for (j in seq_along(inner_parts)) {
        lbl <- extract_leading_label(inner_parts[j])
        blocks[[length(blocks) + 1]] <- list(
          label = if (!is.na(lbl)) lbl else paste0("group_", j),
          text  = inner_parts[j])
      }
      return(blocks)
    }
  }
  
  # FIX (study 173, and 3 other cells with the same shape): "50%+16%" --
  # two bare percentages joined by "+", with nothing else in the cell.
  # Verified against every column in the dataset: all 4 occurrences of this
  # exact shape (study 173's "50%+16%", and three further cells in septic
  # shock/ARDS columns) correspond to a study with a genuine two-group
  # design described in column T (e.g. early/late ICU admission,
  # survivors/non-survivors), so this is treated as two unlabelled
  # subgroups. No text label is available in the cell itself, so generic
  # "group_1"/"group_2" labels are used, consistent with other unlabelled
  # two-way splits elsewhere in this script (e.g. study 150's "+" split).
  pct_plus_pct <- str_match(s, regex(
    "^(\\d+\\.?\\d*)\\s*%\\s*\\+\\s*(\\d+\\.?\\d*)\\s*%\\s*$"))
  
  if (!is.na(pct_plus_pct[1, 1])) {
    return(list(
      list(label = "group_1", text = paste0(pct_plus_pct[1, 2], "%")),
      list(label = "group_2", text = paste0(pct_plus_pct[1, 3], "%"))
    ))
  }
  
  # FIX (study 4, "number of participants" type cells): "n = 575
  # (development cohort n=455, validation cohort n=120)" -- a single overall
  # n=X figure, followed by a bracket containing two or more LABELLED n=Y
  # sub-figures separated by a plain comma (not ";"). This is a different
  # shape from the "overall_lead" pattern above (which requires a fraction
  # N/M and ";"-separated sub-blocks) -- here the overall is a bare n=X (no
  # fraction), and the sub-blocks use "," not ";". Deliberately narrow
  # (requires "label n=number" repeated, comma-joined, inside one bracket)
  # so it only matches this specific recognised shape.
  bare_n_lead <- str_match(s, regex(
    "^n\\s*=\\s*(\\d+)\\s*\\(\\s*([a-zA-Z][a-zA-Z0-9 \\-]*?\\s*n\\s*=\\s*\\d+(?:\\s*,\\s*[a-zA-Z][a-zA-Z0-9 \\-]*?\\s*n\\s*=\\s*\\d+)+)\\s*\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(bare_n_lead[1, 1])) {
    overall_text <- paste0("n=", bare_n_lead[1, 2])
    inner_text   <- bare_n_lead[1, 3]
    inner_parts  <- str_trim(str_split(inner_text, ",")[[1]])
    inner_parts  <- inner_parts[nchar(inner_parts) > 0]
    
    if (length(inner_parts) > 1) {
      blocks <- list(list(label = "overall", text = overall_text))
      for (j in seq_along(inner_parts)) {
        lbl <- extract_leading_label(inner_parts[j])
        blocks[[length(blocks) + 1]] <- list(
          label = if (!is.na(lbl)) lbl else paste0("group_", j),
          text  = inner_parts[j])
      }
      return(blocks)
    }
  }
  
  # FIX (study 150, "number of participants" type cells): "30 (15+15)" -- a
  # total followed by a bracket containing two unlabelled sub-counts joined
  # by "+", which together sum to the total (here 15+15=30). Deliberately
  # narrow (requires EXACTLY two numbers joined by "+" inside the bracket,
  # nothing else) so it does not match unrelated uses of "+" elsewhere
  # (e.g. "56.5 + 7.8" mean/SD notation, which has a decimal and is used as
  # a dispersion indicator rather than a subgroup count split).
  plus_split <- str_match(s, regex(
    "^(\\d+)\\s*\\(\\s*(\\d+)\\s*\\+\\s*(\\d+)\\s*\\)\\s*$"))
  
  if (!is.na(plus_split[1, 1])) {
    return(list(
      list(label = "overall", text = plus_split[1, 2]),
      list(label = "group_1", text = plus_split[1, 3]),
      list(label = "group_2", text = plus_split[1, 4])
    ))
  }
  
  # FIX (study 166, "number of participants" type cells): "182 total (91
  # SOC, 91 CPA)" -- a total followed by a bracket containing two LABELLED
  # sub-counts (number BEFORE its label, not "label n=number" as in the
  # bare_n_lead pattern above) joined by a comma. Deliberately narrow:
  # requires exactly two "number label" segments comma-joined inside one
  # bracket, immediately after a "total" keyword, so it only matches this
  # specific shape and not other comma-separated bracketed content.
  total_labelled <- str_match(s, regex(
    "^(\\d+)\\s*total\\s*\\(\\s*(\\d+)\\s+([a-zA-Z][a-zA-Z0-9\\-]*)\\s*,\\s*(\\d+)\\s+([a-zA-Z][a-zA-Z0-9\\-]*)\\s*\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(total_labelled[1, 1])) {
    return(list(
      list(label = "overall", text = total_labelled[1, 2]),
      list(label = total_labelled[1, 4], text = total_labelled[1, 3]),
      list(label = total_labelled[1, 6], text = total_labelled[1, 5])
    ))
  }
  
  # FIX (study 96, all 4 demographic columns): "Male sex, n (%) 67 (64.4)
  # 69 (65.1) -> 136/210 (64.76%)" -- two unlabelled subgroup value(pct)
  # pairs (corresponding to study 96's bQ/bM arms, though not named as such
  # in this particular cell), followed by an arrow and the combined overall
  # fraction. Verified against every column in the dataset: this exact
  # shape recurs 4 times, all within study 96 (sex, copd, smokers,
  # diabetes), all driven by the same two-arm study design, so this is
  # treated as overall + 2 unlabelled subgroups.
  arrow_pair <- str_match(s, regex(
    "^.*?\\b(\\d+\\.?\\d*)\\s*\\(\\s*([\\d.]+)\\s*\\)\\s+(\\d+\\.?\\d*)\\s*\\(\\s*([\\d.]+)\\s*\\)\\s*->\\s*(\\d+)\\s*/\\s*(\\d+)\\s*\\(\\s*([\\d.]+)\\s*%\\s*\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(arrow_pair[1, 1])) {
    return(list(
      list(label = "overall", text = paste0(arrow_pair[1, 6], "/", arrow_pair[1, 7],
                                            " (", arrow_pair[1, 8], "%)")),
      list(label = "group_1", text = paste0(arrow_pair[1, 2], " (", arrow_pair[1, 3], ")")),
      list(label = "group_2", text = paste0(arrow_pair[1, 4], " (", arrow_pair[1, 5], ")"))
    ))
  }
  
  # FIX (study 96): "Need of mechanical ventilation bQ 65 (63.1) bM 45
  # (42.5)" -- study 96's two treatment arms are named "bQ" and "bM" in
  # column T ("104 in the bQ arm and 106 in the bM arm"), and this same
  # short-label convention recurs directly in several proportion columns.
  # Deliberately scoped to require BOTH "bQ"-style and "bM"-style short
  # labels (1-3 letters then Q, or 1-3 letters then M) each immediately
  # followed by a "value (pct)" pair, so it only matches this specific
  # naming convention and not unrelated text.
  bq_bm_prop <- str_match(s, regex(
    "\\b(b[A-Za-z]{0,2}Q)\\b\\s*(\\d+\\.?\\d*\\s*\\([\\d.]+\\))\\s*\\b(b[A-Za-z]{0,2}M)\\b\\s*(\\d+\\.?\\d*\\s*\\([\\d.]+\\))",
    ignore_case = TRUE))
  
  if (!is.na(bq_bm_prop[1, 1])) {
    return(list(
      list(label = bq_bm_prop[1, 2], text = bq_bm_prop[1, 3]),
      list(label = bq_bm_prop[1, 4], text = bq_bm_prop[1, 5])
    ))
  }
  
  # FIX (study 25, "number of participants" type cells): "total n=2006
  # (including ward patients). sCAP n=204" -- an overall total n=X figure
  # (with an optional parenthetical aside), followed by a full stop and a
  # SHORT named subgroup figure "Label n=Y". Deliberately narrow: requires
  # the cell to consist of EXACTLY these two segments with nothing else, so
  # it does not match unrelated ". "-separated text elsewhere (e.g.
  # "Mean SD (days) n=8.7 (SD 9.6). Median n=6.0 (no IQR)", which describes
  # two different statistics rather than two patient subgroups, or longer
  # multi-segment lists of causative agents).
  overall_plus_named <- str_match(s, regex(
    "^total\\s*n\\s*=\\s*(\\d+)\\s*(?:\\([^)]*\\))?\\s*\\.\\s*([A-Za-z][A-Za-z0-9\\-]*)\\s*n\\s*=\\s*(\\d+)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(overall_plus_named[1, 1])) {
    return(list(
      list(label = "overall", text = paste0("n=", overall_plus_named[1, 2])),
      list(label = overall_plus_named[1, 3],
           text  = paste0("n=", overall_plus_named[1, 4]))
    ))
  }
  
  # FIX (study 114): "No Steroids (n = 125) 21 (16.8). Steroids (n = 83) 28
  # (33.7)" -- two named groups, each written as "Label (n=N) value (pct)",
  # separated by a full stop. Deliberately narrow: requires BOTH segments to
  # follow the exact "Label (n=X) value (pct)" shape, so it only matches
  # this specific recognised format and not other ". "-separated text.
  period_labelled_pair <- str_match(s, regex(
    paste0(
      "^([A-Za-z][A-Za-z \\-]*?)\\s*\\(\\s*n\\s*=\\s*(\\d+)\\s*\\)\\s*(\\d+\\.?\\d*)\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)\\s*\\.\\s*",
      "([A-Za-z][A-Za-z \\-]*?)\\s*\\(\\s*n\\s*=\\s*(\\d+)\\s*\\)\\s*(\\d+\\.?\\d*)\\s*\\(\\s*(\\d+\\.?\\d*)\\s*\\)\\s*$"),
    ignore_case = TRUE))
  
  if (!is.na(period_labelled_pair[1, 1])) {
    return(list(
      list(label = str_trim(period_labelled_pair[1, 2]),
           text  = paste0(period_labelled_pair[1, 4], "/",
                          period_labelled_pair[1, 3], " (",
                          period_labelled_pair[1, 5], "%)")),
      list(label = str_trim(period_labelled_pair[1, 6]),
           text  = paste0(period_labelled_pair[1, 8], "/",
                          period_labelled_pair[1, 7], " (",
                          period_labelled_pair[1, 9], "%)"))
    ))
  }
  
  # FIX (study 166, survival sheet): "inferred 86% SOC and 92% CPA survived
  # hospitalization" -- two named groups, each written as "pct% Label",
  # joined by "and". Deliberately narrow: requires BOTH segments to have a
  # percentage immediately followed by a short label, joined specifically
  # by " and ", so it only matches this recognised shape.
  and_pct_pair <- str_match(s, regex(
    "(\\d+\\.?\\d*)\\s*%\\s+([A-Za-z][A-Za-z\\-]*)\\s+and\\s+(\\d+\\.?\\d*)\\s*%\\s+([A-Za-z][A-Za-z\\-]*)",
    ignore_case = TRUE))
  
  if (!is.na(and_pct_pair[1, 1])) {
    return(list(
      list(label = and_pct_pair[1, 3], text = paste0(and_pct_pair[1, 2], "%")),
      list(label = and_pct_pair[1, 5], text = paste0(and_pct_pair[1, 4], "%"))
    ))
  }
  
  # FIX (study 179): "Hydrocortisone: 25/400 (6.2%)/Placebo: 47/395
  # (11.9%)" -- two named arms separated by "/", each itself a "Label:
  # N/M (pct%)" fragment. Deliberately narrow: requires a text LABEL
  # followed by a colon on BOTH sides of the "/" separator, so an ordinary
  # fraction like "25/400" is never mistaken for this pattern (a bare
  # fraction has no label+colon before the numbers).
  labelled_slash_parts <- str_match(s, regex(
    "^([A-Za-z][A-Za-z0-9 \\-]*?)\\s*:\\s*(\\d+\\s*/\\s*\\d+[^/]*)/\\s*([A-Za-z][A-Za-z0-9 \\-]*?)\\s*:\\s*(\\d+\\s*/\\s*\\d+.*)$",
    ignore_case = TRUE))
  
  if (!is.na(labelled_slash_parts[1, 1])) {
    return(list(
      list(label = str_trim(labelled_slash_parts[1, 2]), text = str_trim(labelled_slash_parts[1, 3])),
      list(label = str_trim(labelled_slash_parts[1, 4]), text = str_trim(labelled_slash_parts[1, 5]))
    ))
  }
  
  # FIX (study 26): "12 / 88 (14%) among patients undergoing lumbar puncture"
  # -- a SINGLE value that is itself describing a named subset of patients
  # (here, "patients undergoing lumbar puncture" corresponds to the "LP"
  # group named in column T: "LP vs no LP groups"), but with no separate
  # overall figure anywhere in the cell. This was previously being treated
  # as the "overall" value, when it should be recognised as a SUBGROUP with
  # no overall reported alongside it.
  #
  # Deliberately scoped to the word "among" specifically (not "in"/"for",
  # which are used in many other grammatical contexts unrelated to
  # subgroups, e.g. "...satisfied the criteria FOR septic shock" -- there
  # "for" names the outcome itself, not a population subset). Also requires
  # there be only ONE such "among" qualifier, no "and"/";"/"vs" elsewhere in
  # the cell (which would indicate two or more groups, handled by the
  # existing splitting logic instead), and that the qualifying phrase itself
  # contains no digit. Verified against every proportion column in the
  # dataset to confirm this only matches the two genuinely relevant cells.
  among_match <- str_locate(s, regex("\\bamong\\s+", ignore_case = TRUE))
  if (!is.na(among_match[1, 1]) &&
      str_detect(substr(s, 1, among_match[1, 1] - 1), "\\d") &&
      str_count(s, regex("\\bamong\\b", ignore_case = TRUE)) == 1 &&
      !str_detect(s, ";") &&
      !str_detect(s, regex("\\bvs\\.?\\b|\\band\\b", ignore_case = TRUE))) {
    
    value_text <- str_trim(substr(s, 1, among_match[1, 1] - 1))
    qualifier  <- str_trim(substr(s, among_match[1, 2] + 1, nchar(s)))
    
    if (!str_detect(qualifier, "\\d") && nchar(qualifier) > 0 && nchar(qualifier) < 60) {
      return(list(list(label = qualifier, text = value_text)))
    }
  }
  
  # Try semicolon split outside parens
  semi_parts <- split_outside_parens(s, ";")
  semi_parts <- semi_parts[nchar(semi_parts) > 0]
  if (length(semi_parts) > 1 && all(sapply(semi_parts, function(p) str_detect(p, "\\d")))) {
    blocks <- lapply(seq_along(semi_parts), function(i) {
      lbl <- extract_leading_label(semi_parts[i])
      list(label = if (!is.na(lbl)) lbl else paste0("group_", i),
           text  = semi_parts[i])
    })
    # FIX (study 199): move any block labelled literally "overall" to the
    # front, so downstream logic (which checks parsed[[1]]$label ==
    # "overall" to decide which block is the overall figure) works
    # regardless of where "overall" appears in the original text -- e.g.
    # "81/185 (44%) overall; 72/156 (46%) indigenous; 9/29 (31%)
    # non-indigenous" has "overall" as a TRAILING label in the first
    # segment, which extract_leading_label's new trailing-label fallback
    # now correctly identifies, but it still needs to be moved to position
    # 1 in case "overall" is not naturally already first.
    overall_pos <- which(sapply(blocks, function(b) tolower(b$label) == "overall"))
    if (length(overall_pos) == 1 && overall_pos != 1) {
      blocks <- c(blocks[overall_pos], blocks[-overall_pos])
    }
    return(blocks)
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

# scalar_val: safely extract a single scalar from anything that might be a
# list, vector, or NULL — always returns exactly one value or NA_character_
# FIX: row[[col]] on a data frame row and vec[i] on a list-type vector can
# return a list rather than a scalar, causing type coercion errors downstream.
scalar_val <- function(x, as_type = "character") {
  if (is.null(x) || length(x) == 0) {
    return(switch(as_type, character = NA_character_, numeric = NA_real_,
                  integer = NA_integer_, NA))
  }
  # Unwrap list
  if (is.list(x)) x <- x[[1]]
  if (is.null(x) || length(x) == 0) {
    return(switch(as_type, character = NA_character_, numeric = NA_real_,
                  integer = NA_integer_, NA))
  }
  switch(as_type,
         character = as.character(x[[1]]),
         numeric   = suppressWarnings(as.numeric(x[[1]])),
         integer   = suppressWarnings(as.integer(x[[1]])),
         x[[1]]
  )
}

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
  
  # FIX (study 95): apply word-to-number conversion before extracting
  # raw_nums, so a spelled-out count like "Nine" is correctly counted among
  # the numbers present in the cell (keeping raw_nums consistent with what
  # split_into_blocks/parse_single_block will themselves find, since both
  # of those also apply this same conversion internally).
  s <- convert_word_numbers(s)
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
  
  # FIX (required for study 26 fix to take effect): previously
  # "length(parsed) == 1" alone caused ANY single-block result to be treated
  # as "overall", even when that single block's label was something else
  # entirely (e.g. a named subgroup like "patients undergoing lumbar
  # puncture" from the new "among" detector above). A single block must now
  # ALSO have the literal label "overall" to be assigned as the overall
  # value; a single block with any other label is correctly treated as a
  # lone subgroup with no separate overall reported.
  if (tolower(parsed[[1]]$label) == "overall") {
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
    # FIX: abs_num_reported should be Y only when the absolute number is
    # present WITHOUT the sample population also being present in the same
    # cell (i.e. a bare count, not part of an N/M fraction). Previously this
    # was Y whenever num was present at all, even alongside pop.
    if (!is.na(ov$num) && is.na(ov$pop)) empty$abs_num_reported  <- "Y"
    if (!is.na(ov$pop))                  empty$abs_frac_reported <- "Y"
    if (!is.na(ov$pct))                  empty$pct_reported      <- "Y"
    if (!is.na(ov$prop_direct))          empty$prop_reported     <- "Y"
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
# flag_duplication: Y if this cell's numbers look like they were wholesale
# copy-pasted from the age column.
#
# FIX (this round -- "Duplication check Y for all?"): the previous version
# flagged Y whenever ANY single number in this cell also appeared ANYWHERE
# in any other proportion column for the same row. This was far too broad:
# the study's total N (column S) legitimately and correctly appears as the
# denominator in many different proportion cells in the same row (e.g.
# study 1's N=86 correctly appears in both "30d mortality: n=31/86" and
# "sex: 15/86"), so almost every row ended up flagged Y, making the flag
# meaningless ("Y for all").
#
# The check this flag was actually designed to catch (study 101, 111) was
# a much more specific problem: the ENTIRE "number of participants" cell
# had been copy-pasted wholesale from the age column (e.g. "60 (49-75)" in
# both). The fix replaces the broad single-number overlap check with a
# whole-cell comparison: are ALL of the numbers in this cell (as a set)
# identical to all of the numbers in the age cell for this row? This only
# fires on genuine wholesale copy-paste, not on a shared population size.
# ---------------------------------------------------------------------------
flag_duplication <- function(cell_nums, row_idx, df, age_col_name) {
  if (length(cell_nums) == 0) return("N")
  if (is.null(age_col_name) || !(age_col_name %in% colnames(df))) return("N")
  age_val <- df[[age_col_name]][row_idx]
  if (is.na(age_val)) return("N")
  age_nums <- as.numeric(str_extract_all(as.character(age_val), "\\d+\\.?\\d*")[[1]])
  if (length(age_nums) == 0) return("N")
  if (setequal(round(cell_nums, 4), round(age_nums, 4))) "Y" else "N"
}

# ---------------------------------------------------------------------------
# build_study_reference: "Surname Year"
#
# FIX: the author column uses TWO different formats in this data:
#   Format A: "Surname F. and Surname2 F2. and ..."     (surname FIRST)
#   Format B: "Surname, Firstname and Surname2, Firstname2 and ..." (comma-separated)
# The previous version always took the LAST word of the first segment as the
# surname, which is correct for Format B (comma already isolates the surname)
# but WRONG for Format A, where the last word is the initial (e.g. "B." from
# "Zhang B."), not the surname "Zhang".
# The fix detects which format is used (presence of a comma before "and")
# and extracts the surname accordingly: text before the comma for Format B,
# or the FIRST word for Format A.
# ---------------------------------------------------------------------------
build_study_reference <- function(author_raw, year_raw) {
  author <- if (isTRUE(!is.na(author_raw)) && length(author_raw) > 0)
    str_trim(as.character(author_raw)) else ""
  year   <- if (isTRUE(!is.na(year_raw))   && length(year_raw)   > 0)
    str_trim(as.character(year_raw))   else ""
  
  if (nchar(author) == 0) {
    surname <- ""
  } else {
    # Isolate the first author: split on " and " (word boundary, case-insens.)
    first_author <- str_trim(str_split(author, regex("\\s+and\\s+", ignore_case = TRUE))[[1]][1])
    
    if (str_detect(first_author, ",")) {
      # Format B: "Surname, Firstname" -> take text before the comma
      surname <- str_trim(str_split(first_author, ",")[[1]][1])
    } else {
      # Format A: "Surname F." or "Surname F.-M." -> take the FIRST word
      surname <- str_trim(str_split(first_author, "\\s+")[[1]][1])
    }
  }
  
  str_trim(paste(surname, year))
}

# =============================================================================
# PART 1: Clean column S (number of participants)
# =============================================================================
# FIX (root cause of "only 4 sheets" + blank study_number + missing year in
# study_reference + all subgroups=0 + all manual calcs missing):
#
# The PREVIOUS fix (check.names = FALSE alone) was INCOMPLETE and did not
# actually solve the problem. Direct testing against the real file confirmed
# that openxlsx::read.xlsx() replaces spaces in column headers with a "."
# via a SEPARATE argument called sep.names, which defaults to "." regardless
# of check.names. So "study number" becomes "study.number", "number of
# participants" becomes "number.of.participants", and so on -- breaking
# every column lookup that uses the original header text with spaces.
#
# check.names in openxlsx controls a different thing (de-duplicating repeated
# names), not space substitution, so it alone never fixed this.
#
# The correct fix is sep.names = " ", which tells openxlsx to keep spaces
# in headers exactly as they appear in the spreadsheet.
df <- read.xlsx(input_path, sheet = 1, check.names = FALSE, sep.names = " ")

# FIX (study 123, 127, 131): openxlsx::read.xlsx() can read a cell that
# contains a clean-looking decimal (e.g. "0.56") and produce a character
# value with binary floating-point noise instead (e.g. "0.56000000000000005"
# or "0.34899999999999998"). This happens at the file-reading stage itself
# and was then propagating into original_observation and every downstream
# calculation. The fix rounds any character value that parses entirely as a
# number (via signif, 10 significant digits -- comfortably more precision
# than any value in this dataset needs) back to a clean decimal string.
# Cells that are not purely numeric (e.g. "14/77 (18.2)") are left untouched,
# since as.numeric() on them returns NA and the original text is kept.
clean_float_noise <- function(x) {
  if (is.na(x) || nchar(x) == 0) return(x)
  num_val <- suppressWarnings(as.numeric(x))
  if (is.na(num_val)) return(x)          # not purely numeric -- leave as-is
  as.character(signif(num_val, 10))
}
df[] <- lapply(df, function(col) {
  if (is.character(col)) vapply(col, clean_float_noise, character(1), USE.NAMES = FALSE)
  else col
})

# Defensive verification: confirm all expected columns are now found.
# If any are still missing, print them so the mismatch can be diagnosed
# immediately rather than silently skipping that column downstream.
expected_cols <- unique(c(
  COL_STUDY_NUM, COL_AUTHORS, COL_YEAR, COL_N_TOTAL, COL_GROUPS,
  COL_GROUP1, COL_GROUP2, COL_SEX_AT, COL_SEX_BD, COL_SEX_BV,
  COL_SMOKE_AV, COL_SMOKE_BW,
  unlist(lapply(PROPORTION_COLS, function(x) x[[2]]))
))
missing_cols <- setdiff(expected_cols, colnames(df))
if (length(missing_cols) > 0) {
  warning("The following expected columns were NOT found in the data: ",
          paste(missing_cols, collapse = ", "),
          ". Check column header spelling/spacing in the source file.")
}

# FIX: vapply enforces a fixed return type (integer length-1) so result is always
# a plain integer vector, never a list, avoiding "list cannot be coerced" errors
n_total_vec <- vapply(
  df[[COL_N_TOTAL]],
  function(x) { v <- clean_n_total(x); if (is.null(v) || is.na(v)) NA_integer_ else as.integer(v) },
  integer(1)
)

# =============================================================================
# PART 2: Build full set of proportion column names for duplication check
# =============================================================================
all_prop_col_names <- unique(unlist(lapply(PROPORTION_COLS, function(x) x[[2]])))
all_prop_col_names <- intersect(all_prop_col_names, colnames(df))

# ---------------------------------------------------------------------------
# compute_checks: manual validation checks for ONE block (overall, subgroup1,
# or subgroup2), each checked against the most appropriate denominator.
#
# FIX (big-picture issue, earlier round): previously there was only ONE set
# of manual checks, computed from a single "primary" value and always
# checked against the OVERALL study total (column S). Each block now gets
# its own full set of checks, with manual_prop/manual_num using the block's
# own reported population (pop) as the denominator when available, falling
# back to the overall study total (n_total) only when the block has no pop
# of its own.
#
# FIX (study 7, this round): pop_equal was comparing a SUBGROUP's own
# population directly against n_total (the whole study's total from column
# S) -- e.g. study 7's "Men" subgroup has pop=176, but n_total=257 (the
# whole study), so 176 was always going to be flagged "N" against 257 even
# though 176 is entirely correct for that subgroup. A subgroup's population
# is not supposed to equal the whole study's total; what it SHOULD equal,
# together with the other subgroup(s), is the overall/study total. pop_equal
# now takes an explicit comparison target (pop_check_against) rather than
# always defaulting to n_total:
#   - for the OVERALL block, this is n_total (the original, correct check)
#   - for a SUBGROUP block, this is the SUM of all subgroup pops (which
#     should equal the overall total) -- passed in by the caller, who has
#     visibility of both subgroups' pop values.
# ---------------------------------------------------------------------------
compute_checks <- function(block_num, block_prop, block_pop, n_total,
                           pop_equal_override = NULL) {
  denom <- if (!is.na(block_pop)) block_pop else n_total
  
  manual_prop <- if (!is.na(block_num) && !is.na(denom) && denom > 0)
    block_num / denom else NA_real_
  prop_equal  <- safe_eq(manual_prop, block_prop)
  
  manual_num  <- if (!is.na(block_prop) && !is.na(denom))
    block_prop * denom else NA_real_
  num_equal   <- safe_eq(manual_num, block_num)
  
  # FIX (study 7, 114 -- corrected properly this time): pop_equal for a
  # SUBGROUP cannot sensibly compare that subgroup's OWN pop against
  # anything directly (a subgroup's pop, e.g. 176 Men, was never going to
  # equal the overall total of 257, nor the sum of both subgroups, 257 --
  # comparing 176 to 257 is comparing the wrong two numbers entirely). The
  # only meaningful consistency check is "do subgroup1_pop + subgroup2_pop
  # together equal the overall/study total?" -- this is computed ONCE by
  # the caller (a single row-level check, not specific to one subgroup) and
  # passed in here as pop_equal_override, applied identically to whichever
  # block is being checked. For the OVERALL block, pop_equal_override is
  # NULL and the original direct block_pop-vs-n_total check is used, which
  # remains correct for that block.
  if (!is.null(pop_equal_override)) {
    pop_equal <- pop_equal_override
  } else {
    pop_equal <- if (!is.na(block_pop)) safe_eq(block_pop, n_total) else NA_character_
  }
  
  list(
    manual_prop = if (!is.na(manual_prop)) round(manual_prop, 6) else NA_real_,
    prop_equal  = prop_equal,
    manual_num  = if (!is.na(manual_num))  round(manual_num,  2) else NA_real_,
    num_equal   = num_equal,
    pop_equal   = pop_equal
  )
}

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
    # FIX: use scalar_val() throughout to guard against list-type returns from
    # row[[col]] (data frame single-row subsetting) and n_total_vec[i]
    # (when vapply produces an integer vector, [i] is safe, but scalar_val
    # provides a consistent interface for all column extractions)
    cell_val  <- scalar_val(combined_vals[i], "character")
    n_total   <- scalar_val(n_total_vec[i],   "numeric")
    
    study_num <- scalar_val(row[[COL_STUDY_NUM]], "character")
    study_ref <- build_study_reference(
      scalar_val(row[[COL_AUTHORS]], "character"),
      scalar_val(row[[COL_YEAR]],    "character"))
    
    # Subgroup detection
    sub_info <- detect_subgroups(
      scalar_val(row[[COL_GROUPS]], "character"),
      scalar_val(row[[COL_GROUP1]], "character"),
      scalar_val(row[[COL_GROUP2]], "character"))
    
    # Parse cell
    parsed <- parse_proportion_cell(cell_val)
    
    # Convenience aliases
    ov_num  <- parsed$overall_num
    ov_prop <- parsed$overall_prop
    ov_pop  <- parsed$overall_pop
    
    # FIX (big-picture issue): compute a FULL, independent set of manual
    # checks for each of the three blocks (overall, subgroup1, subgroup2),
    # each checked against its own most appropriate denominator -- see
    # compute_checks() above for the full rationale.
    #
    # FIX (study 7, 114 -- corrected properly): a subgroup's pop_equal
    # cannot compare that subgroup's OWN pop against anything meaningfully
    # on its own. The only sensible consistency check is whether
    # subgroup1_pop + subgroup2_pop TOGETHER equal the overall population
    # (ov_pop if available, else n_total). This single Y/N result is
    # computed ONCE here and applied identically to both subgroups, since
    # it is a property of the pair as a whole, not of either subgroup
    # individually.
    subgroup_pop_sum <- sum(c(parsed$sub1_pop, parsed$sub2_pop), na.rm = TRUE)
    has_subgroup_pop <- !is.na(parsed$sub1_pop) || !is.na(parsed$sub2_pop)
    pop_compare_target <- if (!is.na(ov_pop)) ov_pop else n_total
    subgroup_pop_equal <- if (has_subgroup_pop)
      safe_eq(subgroup_pop_sum, pop_compare_target)
    else NA_character_
    
    checks_overall <- compute_checks(ov_num,  ov_prop,  ov_pop,  n_total)
    checks_sub1    <- compute_checks(parsed$sub1_num, parsed$sub1_prop, parsed$sub1_pop,
                                     n_total, pop_equal_override = subgroup_pop_equal)
    checks_sub2    <- compute_checks(parsed$sub2_num, parsed$sub2_prop, parsed$sub2_pop,
                                     n_total, pop_equal_override = subgroup_pop_equal)
    
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
    
    # NEW FEATURE: flag_additional_text -- Y if the cell contains
    # substantial descriptive wording beyond the numbers/symbols that were
    # actually parsed (e.g. study 26's "...among patients undergoing lumbar
    # puncture", which qualifies the figure but is not itself captured in
    # any value column). This is distinct from flag_unused_numbers, which
    # only checks for leftover NUMBERS -- this checks for leftover WORDS.
    # All digits and common structural symbols are stripped first; if more
    # than a few characters of text remain, the cell is flagged so a
    # reviewer can check whether that wording changes the figure's meaning.
    text_remainder <- if (!is.na(cell_val)) {
      r <- str_remove_all(as.character(cell_val), "\\d+\\.?\\d*")
      r <- str_remove_all(r, "[()\\[\\]%/;:,.\u2009\u00a0±+~-]")
      str_trim(str_replace_all(r, "\\s+", " "))
    } else ""
    flag_additional_text <- if (nchar(text_remainder) > 3) "Y" else "N"
    
    # Duplication flag (whole-cell comparison against age column only)
    flag_dup <- flag_duplication(raw_nums, i, df, COL_AGE)
    
    # FIX (study 7, 26): has_subgroups/n_subgroups was previously taken
    # ONLY from sub_info (columns T/BB/BC -- the study's overall design),
    # regardless of whether THIS SPECIFIC CELL actually reported a subgroup
    # breakdown. This caused two opposite problems:
    #   - study 7 (icu mortality): column T says "single cohort" (no
    #     contrast), so sub_info said no subgroups -- but the cell itself
    #     names Men/Women subgroups, which were being parsed but then not
    #     flagged.
    #   - study 26 (30d mortality): column T describes FOUR groups
    #     (meningitis vs non-meningitis; LP vs no LP), but THIS cell
    #     ("12/88 (14%) among patients undergoing lumbar puncture") reports
    #     only a single overall figure with no breakdown at all -- so
    #     attaching column T's n_subgroups=4 to this row was misleading.
    #
    # The fix: if the cell itself was successfully split into 2+ blocks by
    # parse_proportion_cell (sub1_label is populated), use that as the
    # authoritative source for THIS row's has_subgroups/n_subgroups. Only
    # when the cell shows no subgroup split of its own do we fall back to
    # the column-T-derived sub_info, since in that situation column T is
    # the only available signal that subgroups exist for the study even
    # though this particular outcome was reported as one overall figure.
    cell_has_subgroups <- !is.na(parsed$sub1_label) || !is.na(parsed$sub1_num) ||
      !is.na(parsed$sub1_prop) || !is.na(parsed$sub1_pop)
    
    if (cell_has_subgroups) {
      n_blocks_in_cell <- length(parsed$parsed_blocks)
      # If an explicit "overall" block exists alongside subgroups, the
      # subgroup count excludes that overall block; otherwise every block
      # found is itself a subgroup.
      has_overall_block <- length(parsed$parsed_blocks) > 0 &&
        tolower(parsed$parsed_blocks[[1]]$label) == "overall"
      final_has_sub <- TRUE
      final_n_sub   <- if (has_overall_block) n_blocks_in_cell - 1L else n_blocks_in_cell
    } else {
      final_has_sub <- sub_info$has_subgroups
      final_n_sub   <- sub_info$n_subgroups
    }
    
    # NEW FEATURE: flag_any -- a single summary column that is "Y" if ANY
    # individual flag in this row is "Y", OR if any of the 15 manual-check
    # *_equal columns is "N". This lets a reviewer filter a sheet down to
    # only the rows that need a closer look, rather than scanning every
    # individual flag/check column separately.
    #
    # NEW FEATURE (this round): any_flag_string -- rather than just a Y/N,
    # this reports WHICH specific column(s) triggered, as a comma-separated
    # string of column headings (e.g. "flag_zero_one, subgroup1_pop_equal").
    # Built from the same two vectors as flag_any, but now NAMED so the
    # triggering column's heading can be reported rather than just a
    # generic Y/N. The names match the actual output column names used in
    # the row list below, so a reviewer can go straight from this string to
    # the relevant column.
    individual_flags <- c(
      flag_not             = parsed$flag_not,
      flag_question        = parsed$flag_question,
      flag_avg_keyword     = parsed$flag_avg_keyword,
      flag_unused_numbers  = flag_unused,
      flag_zero_one        = parsed$flag_zero_one,
      flag_duplication     = flag_dup,
      flag_additional_text = flag_additional_text
    )
    equal_checks <- c(
      overall_prop_equal    = checks_overall$prop_equal,
      overall_num_equal     = checks_overall$num_equal,
      overall_pop_equal     = checks_overall$pop_equal,
      subgroup1_prop_equal  = checks_sub1$prop_equal,
      subgroup1_num_equal   = checks_sub1$num_equal,
      subgroup1_pop_equal   = checks_sub1$pop_equal,
      subgroup2_prop_equal  = checks_sub2$prop_equal,
      subgroup2_num_equal   = checks_sub2$num_equal,
      subgroup2_pop_equal   = checks_sub2$pop_equal
    )
    flag_any <- if (any(individual_flags == "Y", na.rm = TRUE) ||
                    any(equal_checks == "N", na.rm = TRUE)) "Y" else "N"
    
    triggered_flags  <- names(individual_flags)[!is.na(individual_flags) & individual_flags == "Y"]
    triggered_checks <- names(equal_checks)[!is.na(equal_checks) & equal_checks == "N"]
    any_flag_string  <- if (length(triggered_flags) + length(triggered_checks) > 0)
      paste(c(triggered_flags, triggered_checks), collapse = ", ")
    else NA_character_
    
    # NEW FEATURE: flag_excess_subgroups -- "Y" if more than 2 subgroups
    # were identified (n_subgroups > 2). This script only has columns for
    # up to subgroup1/subgroup2, so a study with 3+ subgroups (e.g. column
    # T describing three or more arms) has subgroups that are NOT captured
    # in the subgroup1_*/subgroup2_* columns at all -- this flag highlights
    # those rows so they can be reviewed/extended manually.
    flag_excess_subgroups <- if (!is.na(final_n_sub) && final_n_sub > 2) "Y" else "N"
    
    # NEW FEATURE: all_subgroup_names -- combines subgroup names from every
    # available source: column BB/BC (named study arms), column T (parsed
    # free-text group names), and the individual cell's own subgroup1/2
    # labels (which may name a group the design columns don't, e.g. a
    # cell-specific "Men"/"Women" split).
    all_subgroup_names <- combine_subgroup_names(
      sub_info$names_from_T, sub_info$name_from_BB, sub_info$name_from_BC,
      parsed$sub1_label, parsed$sub2_label)
    
    # Assemble output row in specified column order:
    # study number > study reference > original observation > flags > reported/values > checks
    rows_out[[i]] <- list(
      study_number        = study_num,
      study_reference     = study_ref,
      # FIX: include the original, unparsed cell text so the parsed values
      # can always be checked against the source observation
      original_observation = cell_val,
      # --- FLAGS ---
      flag_any              = flag_any,
      any_flag_string      = any_flag_string,
      flag_not             = parsed$flag_not,
      flag_question        = parsed$flag_question,
      flag_avg_keyword     = parsed$flag_avg_keyword,
      flag_unused_numbers  = flag_unused,
      flag_additional_text = flag_additional_text,
      flag_zero_one        = parsed$flag_zero_one,
      flag_duplication     = flag_dup,
      has_subgroups       = if (isTRUE(final_has_sub)) "Y" else "N",
      n_subgroups         = final_n_sub,
      flag_excess_subgroups = flag_excess_subgroups,
      all_subgroup_names    = all_subgroup_names,
      # --- REPORTED TYPES ---
      abs_num_reported    = parsed$abs_num_reported,
      abs_frac_reported   = parsed$abs_frac_reported,
      pct_reported        = parsed$pct_reported,
      prop_reported       = parsed$prop_reported,
      # --- OVERALL VALUES + CHECKS ---
      overall_num         = ov_num,
      overall_prop        = if (!is.na(ov_prop)) round(ov_prop, 6) else NA_real_,
      overall_pop         = ov_pop,
      overall_manual_prop = checks_overall$manual_prop,
      overall_prop_equal  = checks_overall$prop_equal,
      overall_manual_num  = checks_overall$manual_num,
      overall_num_equal   = checks_overall$num_equal,
      overall_pop_equal   = checks_overall$pop_equal,
      # --- SUBGROUP 1 VALUES + CHECKS ---
      subgroup1_label     = parsed$sub1_label,
      subgroup1_num       = parsed$sub1_num,
      subgroup1_prop      = if (!is.na(parsed$sub1_prop))
        round(parsed$sub1_prop, 6) else NA_real_,
      subgroup1_pop       = parsed$sub1_pop,
      subgroup1_manual_prop = checks_sub1$manual_prop,
      subgroup1_prop_equal  = checks_sub1$prop_equal,
      subgroup1_manual_num  = checks_sub1$manual_num,
      subgroup1_num_equal   = checks_sub1$num_equal,
      subgroup1_pop_equal   = checks_sub1$pop_equal,
      # --- SUBGROUP 2 VALUES + CHECKS ---
      subgroup2_label     = parsed$sub2_label,
      subgroup2_num       = parsed$sub2_num,
      subgroup2_prop      = if (!is.na(parsed$sub2_prop))
        round(parsed$sub2_prop, 6) else NA_real_,
      subgroup2_pop       = parsed$sub2_pop,
      subgroup2_manual_prop = checks_sub2$manual_prop,
      subgroup2_prop_equal  = checks_sub2$prop_equal,
      subgroup2_manual_num  = checks_sub2$manual_num,
      subgroup2_num_equal   = checks_sub2$num_equal,
      subgroup2_pop_equal   = checks_sub2$pop_equal
    )
  }
  
  all_sheet_data[[sheet_name]] <- bind_rows(rows_out)
}

# =============================================================================
# PART 4: Write output workbooks
# =============================================================================

FLAG_COLS  <- c("flag_any","any_flag_string","flag_not","flag_question","flag_avg_keyword","flag_unused_numbers",
                "flag_additional_text","flag_zero_one","flag_duplication","has_subgroups","n_subgroups",
                "flag_excess_subgroups","all_subgroup_names")
CHECK_COLS <- c(
  "overall_manual_prop","overall_prop_equal","overall_manual_num","overall_num_equal","overall_pop_equal",
  "subgroup1_manual_prop","subgroup1_prop_equal","subgroup1_manual_num","subgroup1_num_equal","subgroup1_pop_equal",
  "subgroup2_manual_prop","subgroup2_prop_equal","subgroup2_manual_num","subgroup2_num_equal","subgroup2_pop_equal"
)

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

# =============================================================================
# PART 5: Third workbook -- condensed column subset
#
# NEW FEATURE: a separate workbook containing only the columns requested:
# study number, study reference, original observation, overall_num,
# overall_prop, any_flag_string, all_subgroup_names, subgroup1_label,
# subgroup1_num, subgroup1_prop, subgroup2_label, subgroup2_num,
# subgroup2_prop -- for every sheet, in this exact order.
# =============================================================================
CONDENSED_COLS <- c(
  "study_number", "study_reference", "original_observation",
  "overall_num", "overall_prop",
  "any_flag_string", "all_subgroup_names",
  "subgroup1_label", "subgroup1_num", "subgroup1_prop",
  "subgroup2_label", "subgroup2_num", "subgroup2_prop"
)

build_condensed_sheet_data <- function(sheet_data_list, condensed_cols) {
  lapply(sheet_data_list, function(df_sheet) {
    # Only select columns that actually exist (defensive, in case a future
    # column rename means one of these is briefly missing) -- but warn if
    # any are missing so the gap is visible rather than silent.
    present_cols <- intersect(condensed_cols, colnames(df_sheet))
    missing_cols <- setdiff(condensed_cols, colnames(df_sheet))
    if (length(missing_cols) > 0) {
      warning("Condensed workbook missing expected column(s): ",
              paste(missing_cols, collapse = ", "))
    }
    df_sheet[, present_cols, drop = FALSE]
  })
}

condensed_sheet_data <- build_condensed_sheet_data(all_sheet_data, CONDENSED_COLS)
write_workbook(condensed_sheet_data, output_path_condensed, fill_nr = FALSE)

cat("Done.\n")