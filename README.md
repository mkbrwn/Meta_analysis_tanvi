# Meta-analysis of Hospital Mortality in CAP

This project provides scripts and data for conducting a meta-analysis of hospital mortality in Community-Acquired Pneumonia (CAP) studies.

## Project Structure
- `src/data_clean.r`: Cleans and prepares data, runs meta-analysis with the `metafor` package, and produces a forest plot.
- `src/data_clean_meta_package.r`: Alternative script using the `meta` package's `metaprop` function for meta-analysis and plotting.
- `data/`: Folder for raw data files (e.g., `CAP_mortality.xlsx`).
- `output/`: Folder for output files (e.g., forest plots, summary tables).

## How to Use
1. Place your data file (e.g., `CAP_mortality.xlsx`) in the `data/` folder.
2. Open the R scripts in RStudio or VS Code.
3. Install required R packages if needed:
   - `tidyverse`
   - `readxl`
   - `metafor`
   - `meta`
4. Run the script(s) to clean data, perform meta-analysis, and generate plots.

## Output
- Forest plots and summary statistics will be saved in the `output/` directory.

## Troubleshooting
- Ensure all required R packages are installed.
- Check that your data file matches the expected format (column names, etc.).
- Review script comments and R console error messages for guidance.

## Contact
For further help, contact the project maintainer or consult the script comments for guidance.
