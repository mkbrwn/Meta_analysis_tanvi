# Jesus Is LORD!

# Below from claude after prompt, and brief review. Slightly adapted

# =============================================================================
# Age Data Parsing Script — Version 2
# =============================================================================
# Built from actual data. Handles the following observed formats:
#
# OVERALL (no group split):
#   "42.3 ± 14.7 (Median ± SD)"
#   "Mean age was 60.5 years (SD 16.4)"
#   "≥18 years: years median (IQR) 63.0 (49.0–73.0)"
#   "71 (IQR 64-77)"
#   "74 (SD 6.5)"
#   "adult"
#
# TWO NAMED GROUPS (separated by ; or vs or .):
#   "survivor group (n=49): 70.0 ± 2.0; non-survivor group (n=14): 69.1 ± 4.1"
#   "Median 63 years (SOC) vs 59 years (CPA)"
#   "KP-CAP 68 [56, 80]; SP-CAP 73 [61, 81]"
#   "sirolimus group age years 46.7 +/- 12.1; non-sirolimus group age years 51.5 +/- 16"
#   "mNGS 67y (55-75). CMTs 68y (54-77)"
#
# OVERALL + NAMED GROUPS:
#   "median 50 years (IQR 38-57) overall; 48 (36-55) indigenous; 64 (57-70) non-indigenous"
#
# REPLACE "my_tibble" and "age_column" with your actual object and column names.
# =============================================================================

library(dplyr)
library(stringr)
# Below my additions
library(readxl)
age <- read_xlsx("age.xlsx")

# =============================================================================
# PART 1: parse_age_block()
# =============================================================================
# PURPOSE: Extracts all age-related values from ONE block of text
#          (i.e. text relating to a single group or overall cohort).
# INPUT:   A character string
# OUTPUT:  A named list of 11 values

parse_age_block <- function(text) {
  
  # Initialise all outputs as NA
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
  
  # Clean up the text: trim whitespace, normalise dashes and +/-
  text <- str_trim(text)
  text <- str_replace_all(text, "\u2013|\u2014", "-")  # normalise en/em dash to hyphen
  text <- str_replace_all(text, "\u00b1", "±")          # ensure ± is consistent
  
  # ---------------------------------------------------------------------------
  # CATEGORY
  # ---------------------------------------------------------------------------
  # Looks for:
  #   - "adult" / "Adults>=18y" type phrases
  #   - Age thresholds like ">=18 years", ">18", ">=80 years"
  #   - Age ranges like "18-65y"
  #   - "grouped ages" phrase
  
  cat_match <- str_extract(text, regex(
    paste0(
      "adults?\\s*(>=?\\s*\\d+\\s*y(ears?)?)?|",   # adult, Adults>=18y
      "[>=<]+\\s*\\d+\\s*y(ears?)?|",               # >=18 years, >80y
      "\\d+\\s*-\\s*\\d+\\s*y(ears?)?|",            # 18-65y
      "grouped ages?[^;]*"                           # grouped ages reported
    ),
    ignore_case = TRUE
  ))
  out$category <- if (!is.na(cat_match)) str_trim(cat_match) else NA_character_
  
  # ---------------------------------------------------------------------------
  # TYPE OF AVERAGE
  # ---------------------------------------------------------------------------
  # Observed in data: "mean", "median", "medican" (typo), "average", "mean aged"
  # \\b = word boundary ensures we match whole words only
  
  if (str_detect(text, regex("\\bmedian\\b|\\bmedican\\b", ignore_case = TRUE))) {
    out$type_of_average <- "median"
  } else if (str_detect(text, regex("\\bmean\\b|\\baverage\\b", ignore_case = TRUE))) {
    out$type_of_average <- "mean"
  } else {
    out$type_of_average <- "unknown"
  }
  
  # ---------------------------------------------------------------------------
  # AVERAGE AGE
  # ---------------------------------------------------------------------------
  # Observed patterns:
  #   "mean 60.5", "Mean age was 60.5", "mean +/- SD 75.1",
  #   "median age of 73", "63 (mean)", "Mean 75", "49.5y", "67y"
  #   bare numbers like "52.91", "62.88"
  
  avg <- NA_real_
  
  # Pattern A: number following mean/median/average keyword
  # Handles "mean age was", "median age of", "mean age at admission of", etc.
  m <- str_extract(text, regex(
    "(?:mean\\s*(?:age(?:d|s)?)?(?:\\s*(?:was|of|at admission of))?|median\\s*(?:age(?:\\s*(?:of|,|was))?)?|average\\s*(?:age\\s*was)?)\\s*[~=:]?\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  if (!is.na(m)) avg <- as.numeric(str_extract(m, "\\d+\\.?\\d*$"))
  
  # Pattern B: "63 (mean)" — number BEFORE the word mean in brackets
  if (is.na(avg)) {
    m2 <- str_extract(text, regex("(\\d+\\.?\\d*)\\s*\\(\\s*mean\\s*\\)", ignore_case = TRUE))
    if (!is.na(m2)) avg <- as.numeric(str_extract(m2, "^\\d+\\.?\\d*"))
  }
  
  # Pattern C: bare number (possibly followed by y) before any bracket, ± or end
  # Used as fallback when no keyword found
  if (is.na(avg)) {
    m3 <- str_extract(text, "^[^\\d]*(\\d+\\.?\\d*)\\s*y?\\s*(?:[\\(\\[±\\+]|$)")
    if (!is.na(m3)) avg <- as.numeric(str_extract(m3, "\\d+\\.?\\d*"))
  }
  
  out$average_age <- avg
  
  # ---------------------------------------------------------------------------
  # DISPERSION — checked in priority order: SD > IQR > range > unspecified
  # ---------------------------------------------------------------------------
  
  # --- SD ---
  # Observed: "± 14.7", "+/- SD 75.1 ± 14.0", "(SD 16.4)", "SD 14.6",
  #           "+/- 16.5", "+-14.8", "± 11.9", "61±18", "(±16.1)"
  
  sd_val <- NA_real_
  
  # "SD X" or "(SD X)" — SD keyword followed by number
  m_sd1 <- str_extract(text, regex("\\bSD\\b\\s*(\\d+\\.?\\d*)", ignore_case = TRUE))
  if (!is.na(m_sd1)) sd_val <- as.numeric(str_extract(m_sd1, "\\d+\\.?\\d*$"))
  
  # "± X" or "+/- X" or "+-X" or "(± X)"
  if (is.na(sd_val)) {
    m_sd2 <- str_extract(text, regex("(?:±|\\+\\s*/?\\s*-|\\+-|plus\\s*/?\\s*minus)\\s*(\\d+\\.?\\d*)"))
    if (!is.na(m_sd2)) sd_val <- as.numeric(str_extract(m_sd2, "\\d+\\.?\\d*$"))
  }
  
  # "number±number" e.g. "62±16" with no spaces
  if (is.na(sd_val)) {
    m_sd3 <- str_extract(text, "\\d+\\.?\\d*\\s*±\\s*(\\d+\\.?\\d*)")
    if (!is.na(m_sd3)) sd_val <- as.numeric(str_extract(m_sd3, "\\d+\\.?\\d*$"))
  }
  
  # --- IQR ---
  # Observed: "(IQR 38-57)", "[IQR] 68 [56, 80]", "(Q1 ; Q3) 83 (81; 85)",
  #           "25/75% IQR", "median [IQR] 62 [46; 76]", "IQR 64-77",
  #           "IQR 51 to 71", "(IQR: 20)" — single number (ambiguous, skip pair)
  
  iqr_lo <- NA_real_
  iqr_hi <- NA_real_
  
  # Two-number IQR: "IQR 38-57", "IQR (38, 57)", "Q1 ; Q3" pair
  m_iqr <- str_extract(text, regex(
    "(?:IQR|Q1\\s*[;,]?\\s*Q3|25/?75%?\\s*IQR)\\s*[:\\(\\[]?\\s*(\\d+\\.?\\d*)\\s*(?:[-,;]|to)\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  
  if (!is.na(m_iqr)) {
    nums <- as.numeric(str_extract_all(m_iqr, "\\d+\\.?\\d*")[[1]])
    if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
  } else {
    # Catch "[46; 76]" or "(56, 80)" style where IQR keyword is elsewhere in the cell
    if (str_detect(text, regex("IQR|Q1.*Q3", ignore_case = TRUE))) {
      m_iqr2 <- str_extract(text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[,;-]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
      if (!is.na(m_iqr2)) {
        nums <- as.numeric(str_extract_all(m_iqr2, "\\d+\\.?\\d*")[[1]])
        if (length(nums) >= 2) { iqr_lo <- nums[1]; iqr_hi <- nums[2] }
      }
    }
  }
  
  # --- RANGE ---
  # Observed: "(range 18-95)", "(range 21 to 99)", "range 18 to 101",
  #           "ranged from 14 to 71"
  
  rng_lo <- NA_real_
  rng_hi <- NA_real_
  
  m_rng <- str_extract(text, regex(
    "range[d]?\\s*(?:from)?\\s*(\\d+\\.?\\d*)\\s*(?:[-,]|to)\\s*(\\d+\\.?\\d*)",
    ignore_case = TRUE
  ))
  if (!is.na(m_rng)) {
    nums <- as.numeric(str_extract_all(m_rng, "\\d+\\.?\\d*")[[1]])
    if (length(nums) >= 2) { rng_lo <- nums[1]; rng_hi <- nums[2] }
  }
  
  # --- UNSPECIFIED BRACKETED PAIR ---
  # Observed: "71 (52.5-83.7)", "66 (53-76)", "60 (49-75)", "54.5y (36-73)"
  # Only used if no SD/IQR/range keyword is present
  
  unspec_lo <- NA_real_
  unspec_hi <- NA_real_
  
  has_dispersion_keyword <- str_detect(text, regex("\\bSD\\b|IQR|\\brange\\b|±|\\+/?-", ignore_case = TRUE))
  
  if (!has_dispersion_keyword) {
    m_unspec <- str_extract(text, "[\\(\\[]\\s*(\\d+\\.?\\d*)\\s*[-,;]\\s*(\\d+\\.?\\d*)\\s*[\\)\\]]")
    if (!is.na(m_unspec)) {
      nums <- as.numeric(str_extract_all(m_unspec, "\\d+\\.?\\d*")[[1]])
      if (length(nums) >= 2) { unspec_lo <- nums[1]; unspec_hi <- nums[2] }
    }
  }
  
  # --- Assign dispersion columns based on priority ---
  if (!is.na(sd_val)) {
    out$measure_dispersion <- "SD"
    out$dispersion         <- sd_val
    
  } else if (!is.na(iqr_lo)) {
    out$measure_dispersion <- "IQR"
    out$LQR                <- iqr_lo
    out$UQR                <- iqr_hi
    
  } else if (!is.na(rng_lo)) {
    out$measure_dispersion <- "range"
    out$range_lower        <- rng_lo
    out$range_upper        <- rng_hi
    
  } else if (!is.na(unspec_lo)) {
    out$measure_dispersion <- "unspecified"
    out$lower_unspecified  <- unspec_lo
    out$upper_unspecified  <- unspec_hi
  }
  
  return(out)
}


# =============================================================================
# PART 2: split_groups()
# =============================================================================
# PURPOSE: Determines whether a cell reports overall data, two named groups,
#          or both — and extracts the text block for each.
#
# KEY DESIGN DECISION: Group names are not standardised (e.g. "survivor group",
# "KP-CAP", "sirolimus group", "indigenous"), so we do NOT try to detect them
# by keyword. Instead we:
#   1. Check for an "overall" keyword to identify that segment
#   2. Split remaining text on separators (; / vs / .) into segments
#   3. Treat up to 2 non-overall segments as group_1 and group_2
#   4. Extract the label at the start of each segment as the group name
#
# OUTPUT: A named list with:
#   - overall_reported   (Y / N)
#   - subgroups_reported (Y / N)
#   - group_1_name, group_2_name  (the name label of each group, or NA)
#   - overall_text, group_1_text, group_2_text (text blocks for parsing)

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
  
  text <- str_trim(text)
  text <- str_replace_all(text, "\u2013|\u2014", "-")
  
  # ---------------------------------------------------------------------------
  # STEP 1: Check for an explicit "overall" segment
  # ---------------------------------------------------------------------------
  # Observed: "median 50 years (IQR 38-57) overall; 48 (36-55) indigenous"
  # "overall" appears at the END of a segment in this case
  
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
  # STEP 2: Split remaining text into group segments
  # ---------------------------------------------------------------------------
  # Observed separators: ";" (most common), " vs " , ". " before capital letter
  
  group_segs <- character(0)
  
  if (str_detect(remaining_text, ";")) {
    group_segs <- str_trim(str_split(remaining_text, ";")[[1]])
    group_segs <- group_segs[group_segs != ""]
    
  } else if (str_detect(remaining_text, regex("\\bvs\\.?\\b", ignore_case = TRUE))) {
    group_segs <- str_trim(str_split(remaining_text, regex("\\bvs\\.?\\b", ignore_case = TRUE))[[1]])
    group_segs <- group_segs[group_segs != ""]
    
  } else if (str_detect(remaining_text, "\\.\\s+[A-Z]")) {
    # Split on ". " only when followed by a capital — avoids splitting decimal numbers
    group_segs <- str_trim(str_split(remaining_text, "(?<=\\.)\\s+(?=[A-Z])")[[1]])
    group_segs <- group_segs[group_segs != ""]
  }
  
  # If only one or zero segments remain and no overall found, treat whole cell as overall
  if (length(group_segs) <= 1 && is.na(overall_text)) {
    overall_text <- text
    group_segs   <- character(0)
  }
  
  # ---------------------------------------------------------------------------
  # STEP 3: Extract the group name from the start of each segment
  # ---------------------------------------------------------------------------
  # Strategy: take all non-numeric text at the start, up to the first digit or colon
  # Then clean up trailing noise words (age, years, y, :)
  
  extract_group_name <- function(seg) {
    # Remove sample size notation like "(n=49):" if present
    seg_clean <- str_remove(seg, regex("\\(n\\s*=\\s*\\d+\\)\\s*:?\\s*", ignore_case = TRUE))
    # Extract text before the first digit or colon
    raw_name <- str_extract(seg_clean, "^[^\\d:]+")
    if (is.na(raw_name)) return(NA_character_)
    # Remove trailing words: "age", "years", "year", "y", ":"
    cleaned <- str_remove(raw_name, regex("\\s*(age\\s*)?(years?|y)\\s*$", ignore_case = TRUE))
    cleaned <- str_remove(cleaned, ":\\s*$")
    cleaned <- str_trim(cleaned)
    if (nchar(cleaned) == 0) return(NA_character_)
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
  
  # ---------------------------------------------------------------------------
  # STEP 4: Set reporting flags
  # ---------------------------------------------------------------------------
  
  overall_reported   <- if (!is.na(overall_text))  "Y" else "N"
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
# PURPOSE: Master function — splits the cell into group blocks then parses each.
# INPUT:   A full cell value as a character string
# OUTPUT:  A named list of all output columns

parse_age_cell <- function(text) {
  
  # Step 1: split the cell into group-level text blocks
  groups <- split_groups(text)
  
  # Step 2: parse each block individually
  overall_data <- parse_age_block(groups$overall_text)
  group_1_data <- parse_age_block(groups$group_1_text)
  group_2_data <- parse_age_block(groups$group_2_text)
  
  # Step 3: combine all results into one flat named list with column prefixes
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

# Apply parse_age_cell() to every value in the age column
# Replace "my_tibble" and "age_column" with your actual names
parsed <- lapply(age$age, parse_age_cell)

# Convert list of results into a tibble:
# as.data.frame() converts each result into a one-row dataframe
# bind_rows() stacks all those one-row dataframes into one tibble
parsed_tibble <- bind_rows(
  lapply(parsed, function(x) as.data.frame(x, stringsAsFactors = FALSE))
)

# Attach the new columns to the right of the original tibble
my_tibble_clean <- bind_cols(age, parsed_tibble)

# Preview the result
glimpse(my_tibble_clean)


# =============================================================================
# OPTIONAL: Save to Excel
# =============================================================================

library(openxlsx)

wb <- createWorkbook()
addWorksheet(wb, "cleaned data")
writeData(wb, "cleaned data", my_tibble_clean)
saveWorkbook(wb, "age_data_cleaned_v2.xlsx", overwrite = TRUE)