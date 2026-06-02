#Load packages
library(tidyverse)
library(readxl)
library(meta) # I suspect they used meta package 

# load data
ICU_mortality <- read_excel("data/icu_cardiac_arrest_extracted.xlsx")

#meta-analysis of proportions using meta::metaprop
meta_res <- metaprop(
    event = `ICU cardiac arrest`,
    n = `ICU admissions`, 
    studlab = Study,
    method = "GLMM",           # REML method for estimating between-study variance
    data = ICU_mortality,
    sm = "PLOGIT",           # Logit transformation
    method.ci = "CP",         # Clopper-Pearson confidence intervals
    method.random.ci = "HK",    # Hartung-Knapp adjustment for random pooled effect confidence intervals
    random = TRUE,
    common = FALSE,
    prediction = TRUE,
    backtransf = TRUE)    

summary(meta_res)

forest(
    meta_res, 
    file = "output/icu_mortality_forest_plot.png",
    width = 600
)

#evaluation of DOI and LFK index for publication bias
#remotes::install_github("guido-s/metasens")
library(matasen)
