# DARA-Exam-Analytics
MATLAB toolkit for analysing Spanish PAU examination results from DARA/LECTODARA optical readers, generating statistics, rankings, reports, maps and GIS products.

## Overview

This repository contains a collection of MATLAB scripts developed for the statistical analysis of examination results exported from DARA/LECTODARA optical readers. The workflow was originally developed for the Spanish PAU (University Entrance Examination) but can be adapted to any assessment process based on optical mark recognition (OMR).

## Main features

- Subject-based analysis
- Centre-based analysis
- Statistical summaries
- Composite performance indices
- Automatic Word reports
- Histograms and KDE curves
- Centre rankings
- GIS outputs (Shapefile)
- Final institutional report generation

- ## Input data structure

The workflow is designed to process examination results exported from DARA/LECTODARA optical mark recognition systems.

The input consists of a Microsoft Excel workbook (.xlsx) containing one record per examination and a set of descriptive fields organized by columns.

The scripts currently use the following fields:

| Field | Description |
|---------|---------|
| CENTRO | Educational centre identifier |
| TIPO | Subject code |
| MATERIA | Subject name |
| TIENEEXAMEN | Indicates whether the student took the examination (SI / NO) |
| CALIFICACION | Final score (0–10) |
| COORDENADA_X | Centre X coordinate (optional, used for GIS analyses) |
| COORDENADA_Y | Centre Y coordinate (optional, used for GIS analyses) |

Additional fields may be present and are preserved, although they are not required by the current workflow.

Example:

CENTRO | TIPO | MATERIA | TIENEEXAMEN | CALIFICACION
---------|---------|---------|---------|---------
CENTRO_A | BIO | Biología | SI | 6.75
CENTRO_B | BIO | Biología | SI | 7.80
CENTRO_C | FIS | Física | SI | 5.25

The workflow automatically filters:
- missing scores,
- examinations not taken,
- invalid scores outside the range 0–10.

## Typical workflow

1. Split the master Excel file by subject.
2. Analyse centres within each subject.
3. Generate master statistical tables.
4. Generate global institutional analyses.
5. Generate subject-level reports.
6. Generate the final integrated institutional report.

## Workflow outputs
The toolkit generates:
- Subject-level reports (.docx)
- Global reports (.docx)
- Histograms (.png)
- Vector histograms (.emf)
- Statistical summaries (.xlsx)
- Centre rankings (.xlsx)
- GIS products (.shp)
- Final institutional reports


## Workflow SCRIPST USED

1. `analisis_materia_a_materia.m`
2. `analisis_materia_v4Ab11.m`
3. `analisis_materia_centro_ordinaria3.m`
4. `analisis_pauV2.m`
5. `analisis_materia_v4Aa_PARAEXTRAORDINARIA.m`
6. `analisis_materia_v4_PARAEXTRAORDINARIA.m`
7. `INFORME_FINAL_NUEVO.m`

## Requirements

- MATLAB
- MATLAB Report Generator

## Outputs

- Word reports (.docx)
- Excel summaries (.xlsx)
- Histograms (.png and .emf)
- GIS products (.shp)
- Institutional reports

## Data privacy

This repository does not contain real examination data. Users must provide their own datasets and ensure compliance with local data protection regulations.

## License

MIT License
