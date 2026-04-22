# script to load and clean data for meta-analysis 

#Load packages
library(tidyverse)
library(readxl)
library(meta) # Produces nice plots but does not have REML 

#load data - If you change the brackets to teh file extension of the data sheet the rest of the code should work as is.
dat <- read_excel("data/CAP_mortality.xlsx")


# Clean data
# number_of_patients: The total number of patients included in each study
# hospital_death: The number of hospital deaths, calculated as number_of_patients multiplied by the hospital mortality rate
# prop: The observed proportion of hospital deaths in each study (hospital_death divided by number_of_patients)
dat <- dat %>%
    select(number, year, authors, `number of patients`, `Hospital mortality`)|>  # limit data set to relevant columns    
    mutate( hospital_death = round(as.numeric(`number of patients`) * as.numeric(`Hospital mortality`))) |> # calculate number of hospital deaths
    mutate( study_name = paste0(word(authors, 1), " (", year, ")")) |> 
    mutate( year = as.numeric(year)) |>
    mutate( number_of_patients = as.numeric(`number of patients`)) |>
    rename( proportion_died = `Hospital mortality`) |>
    select(study_name, year, number_of_patients, hospital_death,proportion_died) # select relevant columns for meta-analysis



# Order studies alphabetically by study_name
dat <- dat[order(dat$study_name), ]

#remove if hospital_death is NA
dat <- dat[!is.na(dat$hospital_death), ]

# Meta-analysis of proportions using meta::metaprop
meta_res <- metaprop(
    event = hospital_death,
    n = number_of_patients,
    studlab = study_name,
    data = dat,
    method = "inverse", # Inverse variance method
    sm = "PLOGIT",      # Logit transformation (recommended for heterogeneity)
    random = FALSE, 
    common = FALSE

)

# Print summary
summary(meta_res)

# save as pdf
png( "output/forest_plot.pdf", width = 800, height = 600, dpi = 300)

# Forest plot
     forest(meta_res, sortvar = study_name, 
leftcols = c("studlab", "number_of_patients", "hospital_death", "proportion_died"),
 leftlabs = c("Study", "N", "Deaths", "Proportion"))

# save as pdf
dev.off()


