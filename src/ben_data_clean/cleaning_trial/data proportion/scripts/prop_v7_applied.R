# =============================================================================
# Proportion Data Cleaning Script
# =============================================================================
#
# CONTENTS
# --------
# Header: description ............................................. 1-137
# Libraries ........................................................ 139-142
#
# CONFIGURATION .................................................... 145-195
#   File paths ...................................................... 148-151
#   Column name constants ........................................... 153-166
#   Proportion columns to process ................................... 168-195
#
# HELPERS ........................................................... 198-1006
#   normalise_text() ................................................. 202-209
#   clean_n_total() ................................................... 216-236
#   is_short_label_list() ............................................. 251-260
#   detect_subgroups() ................................................ 272-323
#     Returns has_subgroups, n_subgroups, and names from T/BB/BC ...... 272-323
#   combine_subgroup_names() .......................................... 339-365
#     Merges subgroup names from columns BB, BC, T, and the cell ...... 339-365
#   split_outside_parens() ............................................ 370-387
#   extract_leading_label() ........................................... 395-414
#     Checks leading position first, then trailing position .......... 395-414
#   NUMBER_WORDS / convert_word_numbers() ............................. 416-434
#   fix_typo_fractions() ............................................... 439-447
#   parse_single_block() ............................................... 452-612
#     "All"/"none" special cases ....................................... 457-467
#     Word-to-number conversion + text cleanup ......................... 469-479
#     Pattern: pct% (N/M) ............................................... 481-488
#     Pattern: N/M (pct%) ................................................ 490-499
#     Pattern: N/M (decimal, no %) ....................................... 501-511
#     Pattern: N/M (no percent, no bracket) .............................. 513-519
#     Pattern: N (pct%) ................................................... 521-528
#     Pattern: N (decimal, no %) .......................................... 530-539
#     Pattern: standalone pct% ............................................ 541-547
#     Pattern: N=X ......................................................... 549-554
#     Single number (proportion / count / percentage point) .............. 556-574
#     Two+ numbers: heuristics .............................................. 576-602
#     Fallback: leading integer ............................................. 604-611
#   split_into_blocks() .................................................... 618-825
#     Overall fraction with nested subgroup breakdown (";" inside) ......... 627-648
#     Bare "pct% + pct%" pairs ............................................... 650-658
#     Bare n=X overall with comma-joined subgroup breakdown ................. 660-680
#     Plain "N (A+B)" overall+2-subgroup split ............................... 682-691
#     "N total (A label, B label)" split ...................................... 693-703
#     Two value-pairs joined by an arrow to a combined overall .............. 705-716
#     bQ/bM short-label pairs ................................................. 718-727
#     Total n=X plus a named subgroup figure .................................. 729-739
#     Period-separated "Label (n=X) value (pct)" pairs ....................... 741-758
#     Percentage-before-fraction "and"-joined pairs ........................... 760-769
#     Labelled-slash-separated pairs ........................................... 771-780
#     "among [population]" single-subgroup detection .......................... 782-795
#     Semicolon split outside parens ........................................... 797-811
#     "vs" split ................................................................. 813-822
#     Default: whole cell treated as overall .................................... 824
#   AVG_KEYWORDS_PAT / parse_proportion_cell() ................................... 850-942
#     Master function: parses full cell into overall + subgroups .............. 860-942
#   safe_eq() ...................................................................... 947-954
#   flag_duplication() .............................................................. 960-968
#   build_study_reference() ......................................................... 983-1006
#
# PART 1: Read data, clean column S, verify columns .................................. 1008-1047
#
# PART 2: Build full set of proportion column names, compute_checks() ............... 1048-1085
#
# PART 3: Process each proportion column .............................................. 1085-1290
#   Loop over PROPORTION_COLS .......................................................... 1089-1290
#     Build combined source series ....................................................... 1094-1104
#     Per-row processing .................................................................. 1108-1287
#       Subgroup detection .................................................................. 1121-1125
#       Parse cell ........................................................................... 1127-1128
#       Convenience aliases .................................................................. 1130-1133
#       Manual checks for overall/subgroup1/subgroup2 ....................................... 1135-1146
#       Unused numbers flag .................................................................. 1148-1164
#       Additional-text flag ................................................................. 1166-1171
#       Duplication flag ...................................................................... 1173-1174
#       Cell-aware has_subgroups/n_subgroups ................................................. 1176-1191
#       any_flag_string / flag_any ............................................................ 1193-1220
#       flag_excess_subgroups ................................................................. 1222
#       all_subgroup_names .................................................................... 1224-1226
#       Assemble output row ................................................................... 1228-1287
#
# PART 4: Write output workbooks ......................................... 1293-1359
#   FLAG_COLS / CHECK_COLS ............................................... 1296-1303
#   write_workbook() ...................................................... 1305-1355
#   Save main + NR workbooks .............................................. 1357-1358
#
# PART 5: Third workbook -- condensed column subset ...................... 1361-1389
#   CONDENSED_COLS .......................................................... 1363-1369
#   build_condensed_sheet_data() ............................................ 1371-1384
#   Save condensed workbook .................................................. 1386-1389
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
# OUTPUT WORKBOOKS:
#   proportion_cleaned.xlsx           -- full output, blank cells for NA
#   proportion_cleaned_NR.xlsx        -- full output, "NR" for NA
#   proportion_cleaned_condensed.xlsx -- condensed column subset (see
#                                        CONDENSED_COLS), blank cells for NA
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
output_path     <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v7_add_cleanscript.xlsx"
output_path_nr  <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v7_add_cleanscript_NR.xlsx"
output_path_condensed <- "C:/Users/goddab/OneDrive - University Hospital Southampton NHS Foundation Trust/Desktop/RESEARCH_tanvi_cleaning_trial/proportion data/cleaned data/proportion_cleaned_v7_add_cleanscript_condensed.xlsx"

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
COL_AGE        <- "age"                     # R

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
# Treating every comma as a separator would badly over-split narrative
# cells elsewhere in column T. This scoped heuristic only accepts the
# comma-split when the WHOLE string is short, splits into 2-4 parts, and
# each part reads like a short label (a handful of words, no sentence
# punctuation), so it does not affect longer descriptive cells.
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
CONTRAST_PAT <- regex(
  "\\bvs\\.?\\b|\\bverses?\\b|\\bv/s\\b|\\bv\\b|\\band\\b|;|/|\\bgroup\\b",
  ignore_case = TRUE)

detect_subgroups <- function(groups_val, g1_val, g2_val) {
  g  <- if (isTRUE(!is.na(groups_val)) && length(groups_val) > 0) str_trim(as.character(groups_val)) else ""
  g1 <- if (isTRUE(!is.na(g1_val))    && length(g1_val)    > 0) str_trim(as.character(g1_val))     else ""
  g2 <- if (isTRUE(!is.na(g2_val))    && length(g2_val)    > 0) str_trim(as.character(g2_val))     else ""
  
  # The narrow short-label-list comma case is checked FIRST, before the
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
# combine_subgroup_names: builds the "all_subgroup_names" column by merging
# subgroup names from every available source:
#   - column BB (intervention group) and BC (control group) -- the study's
#     named arms, if directly stated
#   - column T (groups if applicable) -- names parsed out of the free-text
#     description of the study's groups
#   - the individual proportion cell itself -- subgroup1_label/
#     subgroup2_label, which may name a subgroup the study-design columns
#     don't (e.g. "Men"/"Women" appearing only in the icu mortality cell)
#   - cross_sheet_names -- subgroup names found in ANY OTHER sheet's cell
#     for the same study, so a name visible only in one outcome column
#     (e.g. icu mortality) still appears in every sheet's
#     all_subgroup_names for that study
# Every name is tagged with its source in square brackets (e.g. "Men
# [icu mortality]", "development [T]"), and names are deduplicated by
# their bare text (ignoring the tag and case) so the same group named in
# multiple sources is only listed once, keeping the first-seen tag.
# ---------------------------------------------------------------------------
combine_subgroup_names <- function(names_from_T, name_from_BB, name_from_BC,
                                   cell_sub1_label, cell_sub2_label,
                                   cross_sheet_names = character(0)) {
  names_from_T_tagged <- if (length(names_from_T) > 0) paste0(names_from_T, " [T]") else character(0)
  name_from_BB_tagged  <- if (!is.na(name_from_BB)) paste0(name_from_BB, " [BB]") else NA_character_
  name_from_BC_tagged  <- if (!is.na(name_from_BC)) paste0(name_from_BC, " [BC]") else NA_character_
  
  all_names <- c(names_from_T_tagged, name_from_BB_tagged, name_from_BC_tagged,
                 cell_sub1_label, cell_sub2_label, cross_sheet_names)
  all_names <- all_names[!is.na(all_names)]
  all_names <- str_trim(all_names)
  all_names <- all_names[nchar(all_names) > 0]
  if (length(all_names) == 0) return(NA_character_)
  
  # Cosmetic cleanup: if a generic "group_N" placeholder (from a source
  # with no real name available, e.g. an unlabelled "50%+16%" split)
  # appears ALONGSIDE a genuinely named alternative from a different source
  # for the same study (e.g. column T naming "development"/"validation"
  # while the cell itself only produced "group_1"/"group_2"), the generic
  # placeholder is redundant and is dropped. If NO named alternative exists
  # anywhere, the generic placeholder is kept, since it is the only
  # information available.
  is_generic <- str_detect(all_names, regex("^group_\\d+(\\s*\\[[^\\]]+\\])?$", ignore_case = TRUE))
  if (any(is_generic) && any(!is_generic)) {
    all_names <- all_names[!is_generic]
  }
  
  # Deduplicate by name (ignoring the bracketed source tag and case) while
  # preserving the FIRST-seen full tagged string. Two sources naming the
  # same group (e.g. column T and a cell both saying "Men") would otherwise
  # show as two near-identical entries; here only the first occurrence is
  # kept, so the same name is not listed once per source.
  bare_names <- tolower(str_remove(all_names, "\\s*\\[[^\\]]+\\]\\s*$"))
  keep <- !duplicated(bare_names)
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
extract_leading_label <- function(text) {
  m <- str_match(text, "^([A-Za-z][A-Za-z\\s\\-_]*?)\\s*[:/]?\\s*(?=\\d)")
  if (!is.na(m[1,1])) {
    lbl <- str_trim(str_remove(m[1,2], ":$"))
    if (nchar(lbl) > 0 && nchar(lbl) < 40) return(lbl)
  }
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
  
  text <- convert_word_numbers(text)
  if (!str_detect(text, "\\d")) return(out)
  
  s <- fix_typo_fractions(str_trim(text))
  
  s <- str_replace_all(s, "%%+", "%")
  
  s <- str_replace_all(s, "(?<=\\d),(?=\\d{3}(\\D|$))", "")
  
  s <- str_remove(s, regex("\\s*(?:both\\s+)?n\\s*,?\\s*\\(\\s*%\\s*\\)\\s*$", ignore_case = TRUE))
  s <- str_trim(s)
  
  m <- str_match(s, "(\\d+\\.?\\d*)\\s*%\\s*\\(\\s*(\\d+)\\s*/\\s*(\\d+)\\s*\\)")
  if (!is.na(m[1, 1])) {
    out$pct  <- as.numeric(m[1, 2])
    out$prop <- out$pct / 100
    out$num  <- as.numeric(m[1, 3])
    out$pop  <- as.numeric(m[1, 4])
    return(out)
  }
  
  # Pattern: N/M (pct%)  e.g. "31/86 (36.0%)" or "n=31/86 (36%)"
  # Also matches when % is separated by extra spaces e.g. "722/1166  61.9%"
  m <- str_match(s, "(\\d+)\\s*/\\s*(\\d+)\\s*[\\(\\s]*(\\d+\\.?\\d*)\\s*%")
  if (!is.na(m[1,1])) {
    out$num  <- as.numeric(m[1,2])
    out$pop  <- as.numeric(m[1,3])
    out$pct  <- as.numeric(m[1,4])
    out$prop <- out$pct / 100
    return(out)
  }
  
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
  
  text <- convert_word_numbers(text)
  if (!str_detect(text, "\\d"))
    return(list(list(label = "overall", text = text %||% "")))
  
  s <- str_replace_all(str_trim(text), "\\s+", " ")
  
  overall_lead <- str_match(s, regex(
    "^(?:total|overall|n)\\s*n?\\s*[=:]?\\s*(\\d+)\\s*/\\s*(\\d+)\\s*(?:\\(([\\d.]+)%\\)\\s*)?\\((.*;.*)\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(overall_lead[1, 1])) {
    overall_pct  <- overall_lead[1, 4]
    overall_text <- if (!is.na(overall_pct))
      paste0(overall_lead[1, 2], "/", overall_lead[1, 3], " (", overall_pct, "%)")
    else
      paste0(overall_lead[1, 2], "/", overall_lead[1, 3])
    inner_text   <- overall_lead[1, 5]
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
  
  pct_plus_pct <- str_match(s, regex(
    "^(\\d+\\.?\\d*)\\s*%\\s*\\+\\s*(\\d+\\.?\\d*)\\s*%\\s*$"))
  
  if (!is.na(pct_plus_pct[1, 1])) {
    return(list(
      list(label = "group_1", text = paste0(pct_plus_pct[1, 2], "%")),
      list(label = "group_2", text = paste0(pct_plus_pct[1, 3], "%"))
    ))
  }
  
  ci_pair <- str_match(s, regex(
    "^([A-Za-z][A-Za-z \\-]*?)\\s*%\\s*\\(\\s*95%\\s*CI\\s*\\)\\s*(\\d+\\.?\\d*)\\s*\\([^)]*\\)\\s*([A-Za-z][A-Za-z \\-]*?)\\s*%\\s*\\(\\s*95%\\s*CI\\s*\\)\\s*(\\d+\\.?\\d*)\\s*\\([^)]*\\)\\s*$",
    ignore_case = TRUE))
  
  if (!is.na(ci_pair[1, 1])) {
    return(list(
      list(label = str_trim(ci_pair[1, 2]), text = paste0(ci_pair[1, 3], "%")),
      list(label = str_trim(ci_pair[1, 4]), text = paste0(ci_pair[1, 5], "%"))
    ))
  }
  
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
  
  plus_split <- str_match(s, regex(
    "^(\\d+)\\s*\\(\\s*(\\d+)\\s*\\+\\s*(\\d+)\\s*\\)\\s*$"))
  
  if (!is.na(plus_split[1, 1])) {
    return(list(
      list(label = "overall", text = plus_split[1, 2]),
      list(label = "group_1", text = plus_split[1, 3]),
      list(label = "group_2", text = plus_split[1, 4])
    ))
  }
  
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
  
  bq_bm_prop <- str_match(s, regex(
    "\\b(b[A-Za-z]{0,2}Q)\\b\\s*(\\d+\\.?\\d*\\s*\\([\\d.]+\\))\\s*\\b(b[A-Za-z]{0,2}M)\\b\\s*(\\d+\\.?\\d*\\s*\\([\\d.]+\\))",
    ignore_case = TRUE))
  
  if (!is.na(bq_bm_prop[1, 1])) {
    return(list(
      list(label = bq_bm_prop[1, 2], text = bq_bm_prop[1, 3]),
      list(label = bq_bm_prop[1, 4], text = bq_bm_prop[1, 5])
    ))
  }
  
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
  
  and_pct_pair <- str_match(s, regex(
    "(\\d+\\.?\\d*)\\s*%\\s+([A-Za-z][A-Za-z\\-]*)\\s+and\\s+(\\d+\\.?\\d*)\\s*%\\s+([A-Za-z][A-Za-z\\-]*)",
    ignore_case = TRUE))
  
  if (!is.na(and_pct_pair[1, 1])) {
    return(list(
      list(label = and_pct_pair[1, 3], text = paste0(and_pct_pair[1, 2], "%")),
      list(label = and_pct_pair[1, 5], text = paste0(and_pct_pair[1, 4], "%"))
    ))
  }
  
  labelled_slash_parts <- str_match(s, regex(
    "^([A-Za-z][A-Za-z0-9 \\-]*?)\\s*:\\s*(\\d+\\s*/\\s*\\d+[^/]*)/\\s*([A-Za-z][A-Za-z0-9 \\-]*?)\\s*:\\s*(\\d+\\s*/\\s*\\d+.*)$",
    ignore_case = TRUE))
  
  if (!is.na(labelled_slash_parts[1, 1])) {
    return(list(
      list(label = str_trim(labelled_slash_parts[1, 2]), text = str_trim(labelled_slash_parts[1, 3])),
      list(label = str_trim(labelled_slash_parts[1, 4]), text = str_trim(labelled_slash_parts[1, 5]))
    ))
  }
  
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
# list, vector, or NULL — always returns exactly one value or NA_character_.
# row[[col]] on a data frame row and vec[i] on a list-type vector can
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
    # abs_num_reported is Y only when an absolute number is present
    # WITHOUT the sample population also being present in the same cell
    # (i.e. a bare count, not part of an N/M fraction).
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
# The author column uses two different formats:
#   Format A: "Surname F. and Surname2 F2. and ..."     (surname FIRST)
#   Format B: "Surname, Firstname and Surname2, Firstname2 and ..." (comma-separated)
# Taking the LAST word of the first segment as the surname is correct for
# Format B (the comma already isolates the surname) but wrong for Format A,
# where the last word is the initial (e.g. "B." from "Zhang B."), not the
# surname "Zhang". The function detects which format is used (presence of
# a comma before "and") and extracts the surname accordingly: text before
# the comma for Format B, or the first word for Format A.
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
df <- read.xlsx(input_path, sheet = 1, check.names = FALSE, sep.names = " ")

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

# Defensive verification: confirm all expected columns are found.
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

# vapply enforces a fixed return type (integer length-1) so the result is
# always a plain integer vector, never a list.
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
compute_checks <- function(block_num, block_prop, block_pop, n_total,
                           pop_equal_override = NULL,
                           allow_fallback_to_total = TRUE) {
  # The denominator for manual_prop/manual_num should be the block's own
  # population when available. Falling back to n_total (the whole study's
  # total) only makes sense for the OVERALL block -- for a SUBGROUP with no
  # population of its own, n_total is a different, larger population than
  # what the subgroup's proportion was actually computed from, so using it
  # produces a manual_num far bigger than the subgroup could ever contain.
  # allow_fallback_to_total = FALSE (used for subgroup calls) keeps the
  # denominator as NA in that situation, so manual_prop/manual_num are
  # correctly left blank rather than showing a misleading inflated figure.
  denom <- if (!is.na(block_pop)) block_pop
  else if (allow_fallback_to_total) n_total
  else NA_real_
  
  manual_prop <- if (!is.na(block_num) && !is.na(denom) && denom > 0)
    block_num / denom else NA_real_
  prop_equal  <- safe_eq(manual_prop, block_prop)
  
  manual_num  <- if (!is.na(block_prop) && !is.na(denom))
    block_prop * denom else NA_real_
  num_equal   <- safe_eq(manual_num, block_num)
  
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
# PART 2b: Collect cell-level subgroup names from every sheet, per study
#
# A subgroup name may only be visible in one specific sheet's cell text
# (e.g. "Men"/"Women" appearing only in the icu mortality cell, with no
# equivalent in columns T/BB/BC or any other outcome column). This first
# pass loops through every sheet and every row, extracts any subgroup
# names found in that cell, and records them against the study number so
# that every sheet's all_subgroup_names column can show the full set of
# names identified anywhere for that study, each tagged with the column
# it came from.
# =============================================================================
cell_level_names_by_study <- new.env()

for (col_def in PROPORTION_COLS) {
  sheet_src_cols <- col_def[[2]]
  existing_cols  <- intersect(sheet_src_cols, colnames(df))
  if (length(existing_cols) == 0) next
  tag <- existing_cols[1]
  
  combined_vals_pass1 <- apply(df[, existing_cols, drop = FALSE], 1, function(row) {
    v <- row[!is.na(row) & row != ""]
    if (length(v) == 0) NA_character_ else as.character(v[1])
  })
  
  for (i in seq_len(nrow(df))) {
    study_key <- scalar_val(df[i, ][[COL_STUDY_NUM]], "character")
    if (is.na(study_key) || nchar(study_key) == 0) next
    
    cell_val_pass1 <- scalar_val(combined_vals_pass1[i], "character")
    parsed_pass1   <- parse_proportion_cell(cell_val_pass1)
    
    found <- c(parsed_pass1$sub1_label, parsed_pass1$sub2_label)
    found <- found[!is.na(found)]
    found <- found[!str_detect(found, regex("^group_\\d+$", ignore_case = TRUE))]
    if (length(found) == 0) next
    
    tagged <- paste0(found, " [", tag, "]")
    existing <- if (exists(study_key, envir = cell_level_names_by_study, inherits = FALSE))
      get(study_key, envir = cell_level_names_by_study) else character(0)
    assign(study_key, c(existing, tagged), envir = cell_level_names_by_study)
  }
}

get_cell_level_names_for_study <- function(study_key) {
  if (is.na(study_key) || !exists(study_key, envir = cell_level_names_by_study, inherits = FALSE))
    return(character(0))
  get(study_key, envir = cell_level_names_by_study)
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
    # scalar_val() guards against list-type returns from row[[col]]
    # (data frame single-row subsetting), providing a consistent interface
    # for all column extractions.
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
    
    subgroup_pop_sum <- sum(c(parsed$sub1_pop, parsed$sub2_pop), na.rm = TRUE)
    has_subgroup_pop <- !is.na(parsed$sub1_pop) || !is.na(parsed$sub2_pop)
    pop_compare_target <- if (!is.na(ov_pop)) ov_pop else n_total
    subgroup_pop_equal <- if (has_subgroup_pop)
      safe_eq(subgroup_pop_sum, pop_compare_target)
    else NA_character_
    
    checks_overall <- compute_checks(ov_num,  ov_prop,  ov_pop,  n_total)
    checks_sub1    <- compute_checks(parsed$sub1_num, parsed$sub1_prop, parsed$sub1_pop,
                                     n_total, pop_equal_override = subgroup_pop_equal,
                                     allow_fallback_to_total = FALSE)
    checks_sub2    <- compute_checks(parsed$sub2_num, parsed$sub2_prop, parsed$sub2_pop,
                                     n_total, pop_equal_override = subgroup_pop_equal,
                                     allow_fallback_to_total = FALSE)
    
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
    
    text_remainder <- if (!is.na(cell_val)) {
      r <- str_remove_all(as.character(cell_val), "\\d+\\.?\\d*")
      r <- str_remove_all(r, "[()\\[\\]%/;:,.\u2009\u00a0±+~-]")
      str_trim(str_replace_all(r, "\\s+", " "))
    } else ""
    flag_additional_text <- if (nchar(text_remainder) > 3) "Y" else "N"
    
    # Duplication flag (whole-cell comparison against age column only)
    flag_dup <- flag_duplication(raw_nums, i, df, COL_AGE)
    
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
    
    flag_excess_subgroups <- if (!is.na(final_n_sub) && final_n_sub > 2) "Y" else "N"
    
    # Tag this sheet's own cell-level subgroup labels with the source
    # column, consistent with how cross-sheet names are tagged.
    current_sheet_tag <- existing_cols[1]
    sub1_label_tagged <- if (!is.na(parsed$sub1_label) &&
                             !str_detect(parsed$sub1_label, regex("^group_\\d+$", ignore_case = TRUE)))
      paste0(parsed$sub1_label, " [", current_sheet_tag, "]")
    else parsed$sub1_label
    sub2_label_tagged <- if (!is.na(parsed$sub2_label) &&
                             !str_detect(parsed$sub2_label, regex("^group_\\d+$", ignore_case = TRUE)))
      paste0(parsed$sub2_label, " [", current_sheet_tag, "]")
    else parsed$sub2_label
    
    cross_sheet_names <- get_cell_level_names_for_study(study_num)
    
    all_subgroup_names <- combine_subgroup_names(
      sub_info$names_from_T, sub_info$name_from_BB, sub_info$name_from_BC,
      sub1_label_tagged, sub2_label_tagged, cross_sheet_names)
    
    # Assemble output row in specified column order:
    # study number > study reference > original observation > flags > reported/values > checks
    rows_out[[i]] <- list(
      study_number        = study_num,
      study_reference     = study_ref,
      # The original, unparsed cell text, so the parsed values can always
      # be checked against the source observation
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