# script to load and clean data for meta-analysis 

#Load packages
library(tidyverse)
library(readxl)
library(metafor)

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
    

# Pooling proportions of hospital deaths
# Calculate observed proportions and pool using rma (random-effects meta-analysis)
# Use escalc to compute effect sizes (proportions)
escalc_dat <- escalc(measure = "PLO", xi = hospital_death, ni = number_of_patients, data = dat)

# Random-effects meta-analysis of proportions
res <- rma.glmm(
                measure = "PLO",
                yi,
                 vi, 
                 data = escalc_dat, 
                 method = "REML")


# Print summary
summary(res)

# Plot forest plot with ilab columns for number_of_patients, hospital_death, and prop
ilab_data <- cbind(dat$number_of_patients, dat$hospital_death, round(dat$proportion_died, 3))
colnames(ilab_data) <- c("N", "Deaths", "Proportion")

forest.rma(
    res, 
    slab = dat$study_name,
    showweights = TRUE,
    ilab = ilab_data,
    ilab.xpos = c(-7, -5.75, -4.5)
)

# Add column headers for ilab columns (aligned with main headings)
op <- par(cex = 1)
text(c(-7, -5.75, -4.5), par("usr")[4], c("N", "Deaths", "Proportion"), pos = 1, font = 2)
par(op)

# Add summary of heterogeneity below the plot
het_text <- paste0(
    "Heterogeneity: Q = ", round(res$QE, 2),
    ", p = ", format.pval(res$QEp, digits = 2),
    ", I² = ", round(res$I2, 1), "%",
    ", tau² = ", round(res$tau2, 3)
)
text(x = -8.8, y = -2, labels = het_text, pos = 4, cex = 0.9)

# Add description of the overall effect below the plot
overall_text <- paste0(
    "Overall effect (logit scale): ", round(res$b, 2),
    " [", round(res$ci.lb, 2), ", ", round(res$ci.ub, 2), "]"
)
text(x = -8.8, y = -1, labels = overall_text, pos = 4, cex = 0.9)
