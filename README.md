# Cancer Cell Data Cleaning Pipeline

This project is a Python-based data cleaning and filtering program developed in collaboration with the Kosoff Oncology Lab. It processes large-scale experimental datasets from over one million cancer cells, identifies and handles spatially redundant measurements, and standardizes output formats for analysis.

The core functionality involves scanning through microscope-generated Excel spreadsheets, identifying overlapping or near-duplicate cell entries based on X/Y spatial coordinates and layer depth, and extracting only the most representative measurements. This helps reduce noise and redundancy in cancer cell profiling experiments.

---

## Project Background

This tool was developed to support the Kosoff Lab's cancer imaging and cell characterization studies. Given the massive volume of spatial data per experiment, manual filtering was unfeasible. This automated pipeline has been iteratively refined based on ongoing lab feedback, helping improve the accuracy and consistency of analysis-ready data.

---

## Features

- Batch processes all `.xlsx` files in the working directory
- Filters microscope data based on spatial proximity (`CentreX [µm]`, `CentreY [µm]`) and vertical layer (`ND.Z`)
- Detects overlapping entries and retains only median duplicates
- Separates "cleansed" datasets from "excised" (removed) entries
- Generates a count of entries per Z-layer (1–101) for distribution analysis
- Outputs results to new `.xlsx` files (`Cleansed` and `Excised`)

---

## Technologies Used

- **Python**
- **pandas** – data manipulation
- **NumPy** – numerical operations
- **glob** / **os** – batch file I/O

---

## Input Format

Each `.xlsx` file is expected to contain:

- At least 24 metadata rows at the bottom (which are separately extracted)
- Cell data with at minimum the following columns:
  - `Item`
  - `ND.Z`
  - `CentreX [µm]`
  - `CentreY [µm]`

---

## How It Works

1. **Metadata Extraction**  
   The bottom 24 rows of each `.xlsx` file are extracted and stored for reference.

2. **Data Filtering**  
   The script:
   - Iterates through each row
   - Searches for neighboring rows within a 5-layer range
   - Computes spatial distance (threshold: 6.19 µm)
   - If duplicates are found, retains only the median row

3. **Aggregation**  
   - Remaining clean rows + median duplicates = final output
   - Also generates a per-layer count of surviving entries (1–101)

4. **Output Files**  
   - `<filename> Cleansed.xlsx` → clean dataset with counts
   - `<filename> Excised.xlsx` → all removed duplicates (for auditability)

---
