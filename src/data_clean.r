# script to load and clean data for meta-analysis 

#Load packages
library(tidyverse)
library(readxl)
library(meta) # I suspect they used meta package 

#load data - If you change the brackets to teh file extension of the data sheet the rest of the code should work as is.
dat <- read_excel("data/CAP_mortality.xlsx")


# Clean data
# number_of_patients: The total number of patients included in each study
# hospital_death: The number of hospital deaths, calculated as number_of_patients multiplied by the hospital mortality rate
# prop: The observed proportion of hospital deaths in each study (hospital_death divided by number_of_patients)
dat <- dat %>%
    select(number, year, authors, `number of patients`, `Hospital mortality`)|>  # limit data set to relevant columns    
    mutate( hospital_death = as.numeric(`number of patients`) * as.numeric(`Hospital mortality`)) |> # calculate number of hospital deaths
    mutate( study_name = paste0(word(authors, 1), " (", year, ")")) |> 
    mutate( year = as.numeric(year)) |>
    mutate( number_of_patients = as.numeric(`number of patients`)) |>
    rename( proportion_died = `Hospital mortality`) |>
    select(study_name, year, number_of_patients, hospital_death,proportion_died) # select relevant columns for meta-analysis

# Order studies alphabetically by study_name
dat <- dat[order(dat$study_name), ]

#remove if hospital_death is NA
dat <- dat[!is.na(dat$hospital_death), ] 

#make numeric 
dat <- dat %>%
    mutate(
        hospital_death = round(as.numeric(hospital_death)),
        number_of_patients = round(as.numeric(number_of_patients)),
        proportion_died = as.numeric(proportion_died)
    )

# Meta-analysis of proportions using meta::metaprop
res <- metaprop(
    event = hospital_death,
    n = number_of_patients,
    studlab = dat$study_name,
    method = "GLMM",           # REML method for estimating between-study variance
    data = dat,
    sm = "PLOGIT",           # Logit transformation
    method.ci = "CP",        # Clopper-Pearson confidence intervals
    random = TRUE,
    common = TRUE,
    prediction = TRUE)

# Print summary
summary(res)

#produce forest plot and save as png
forest(
    res, 
    file = "output/forest_plot.png",
    width = 600
)
