# ddPCR HDR Analysis Dashboard

A Streamlit web application for automated analysis of digital droplet PCR (ddPCR) HDR (homology-directed repair) data from a Bio-Rad QX200 Droplet Reader. The app accepts two Excel inputs, calculates per-sample HDR rates and replicate statistics, and produces a formatted, downloadable Excel report with embedded figures.

## Overview

This tool is designed to streamline the analysis of ddPCR-based HDR validation experiments. Given a raw QX200 output file and a sample key mapping well positions to sample names, the app computes HDR allele frequencies, flags low-quality replicates, and exports a report ready for downstream use in tools like Benchling.

## Inputs

### QX200 File
The raw Excel output from the Bio-Rad QX200 Droplet Reader. The following columns are required:

| Column | Description |
|--------|-------------|
| `Well` | Plate well identifier |
| `Target` | Assay target name (FAM or HEX) |
| `Conc(copies/µL)` | Droplet digital PCR concentration |
| `Accepted Droplets` | Total accepted droplet count |
| `Positives` | Positive droplet count |
| `Negatives` | Negative droplet count |

### Sample Key
An Excel file mapping well positions to sample metadata. Required columns:

| Column | Description |
|--------|-------------|
| `Well` | Plate well identifier (must match QX200 file) |
| `Name` | Sample name (see naming conventions below) |
| `Sample Entity Link` | Optional — Benchling sample entity link |
| `Analytical Control Link` | Optional — Benchling analytical control link |

A template (`SampleKeyTemplate.xlsx`) is available for download directly from the app.

#### Sample Naming Conventions
Sample names must be underscore-delimited, where the **last field indicates the replicate**. The app groups replicates by stripping this final field. Examples of valid replicate pairs:

- `LKP23059_MOI40k_Rep1` and `LKP23059_MOI40k_Rep2`
- `LKP23059_1` and `LKP23059_2`
- `LKP23059_4.58RNP_40kMOI_A` and `LKP23059_4.58RNP_40kMOI_B`

The number of fields does not matter as long as the replicate identifier is last.

## Parameters

The following analysis parameters can be adjusted in the app before uploading files:

| Parameter | Default | Description |
|-----------|---------|-------------|
| Output file name | `25BCPXXX_AnalyzedResults.xlsx` | Name of the output Excel report |
| FAM target name | `CCR5` | Name of the FAM channel target as it appears in the QX200 file |
| HEX target name | `CCRL2` | Name of the HEX channel target as it appears in the QX200 file |
| CV threshold | `5` | Replicate CV (%) above which results are flagged in red |

## Analysis

For each sample group the app calculates:

- **HDR (%)** — per-replicate HDR allele frequency: `100 × FAM Conc / HEX Conc`
- **Avg. HDR (%)** — mean HDR across replicates
- **Replicate CV** — coefficient of variation across replicates

### Pass/Fail Criteria
A sample group is marked **Pass** only if all of the following are met:

- Mean FAM concentration ≥ 20 copies/µL
- Mean HEX concentration ≥ 20 copies/µL
- Mean accepted droplets ≥ 10,000
- Replicate CV ≤ CV threshold (default 20%)

## Outputs

### Excel Report
A single `.xlsx` file written to the `Results/` directory containing:

- **Main sheet** — full per-well data table with HDR values, colored by:
  - `Avg. HDR(%)` column — red-to-green pastel heatmap
  - `Replicate CV` column — red highlighting for values exceeding the CV threshold
  - Embedded bar graph of average HDR values per sample group
  - Embedded droplet scatter plot
- **Benchling Queryable Output sheet** — one row per sample group with the following fields, formatted for direct import into Benchling:

| Column | Description |
|--------|-------------|
| Sample ID (Text) | Sample group name |
| Sample Entity Link | Benchling entity link |
| Analytical Control Link | Benchling control link |
| % Targeted Integration | Average HDR (%) |
| Copies/uL (FAM) | Mean FAM concentration |
| Copies/uL (HEX) | Mean HEX concentration |
| Droplet Number | Mean accepted droplets |
| %CV (integration) | Replicate CV |
| Pass/Fail | QC pass/fail call |
| Primer Target (FAM) | FAM target name |
| Primer Target (HEX) | HEX target name |

### Figures
Both figures are embedded in the Excel report and available for individual download from the app:

- **HDR Bar Graph** — per-sample-group average HDR values
- **Droplet Stats Scatter Plot** — positive vs. negative droplets colored by target

## Usage

### Installation

```bash
conda env create -f streamlit_dashboard.yml
conda activate streamlit_dashboard
```

### Running the App

```bash
streamlit run app.py
```

Then open the URL printed in the terminal (typically `http://localhost:8501`) in your browser.

## Requirements

See `streamlit_dashboard.yml` for the full environment specification. Key dependencies:

- Python 3.13
- pandas 3.0.2
- streamlit
- matplotlib
- seaborn
- openpyxl
- numpy
