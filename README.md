# Universal Precision Validation — 5x5 Format

This is a separate Precision application.

## Population logic

### INTRA
The supplied Precision template uses:
- 5 replicates per run
- 5 Intra runs
- 25 Intra observations in total

### INTER
- INTER Run 1 = first 5 INTRA results
- INTER Run 2 = first uploaded INTER PDF
- INTER Run 3 = second uploaded INTER PDF
- INTER Run 4 = third uploaded INTER PDF
- Continue in upload order

No additional INTER run is populated unless the corresponding INTER PDF is uploaded.

## Output
- Keeps the uploaded Excel workbook as the starting template.
- Generates analyte-wise workbooks.
- Rebuilds N, Mean, SD and CV% formulas.
- Creates one ZIP containing all analyte workbooks.

## Important
The PDF extraction logic currently supports the same PR INTRA/INTER result-record format used by the existing Precision application. Template population is based on the 5-replicate × 5-run structure.
