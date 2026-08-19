# DARA-Exam-Analytics

MATLAB toolkit for analysing examination results obtained from either DARA/LECTODARA optical mark recognition systems or optical readers, generating statistics, rankings, reports, maps and GIS products.

The workflow was originally developed for the Spanish University Entrance Examination (PAU) but can be adapted to any large-scale assessment process based on optical mark recognition (OMR).

## Overview

This repository contains a collection of MATLAB scripts developed for the statistical analysis of examination results exported from DARA/LECTODARA optical readers. The workflow was originally developed for the Spanish PAU (University Entrance Examination) but can be adapted to any assessment process based on optical mark recognition (OMR).

---

# Main Features

- Subject-based analysis
- Centre-based analysis
- Composite performance indices
- Statistical summaries
- Automatic Word reports
- Subject rankings
- Centre rankings
- Histograms and density curves
- GIS outputs (Shapefile)
- Final institutional report generation

---

# Required MATLAB Toolboxes

Mandatory:

- MATLAB
- MATLAB Report Generator

Optional:

- Statistics and Machine Learning Toolbox
- Mapping Toolbox

---

# Input Data Structure

The workflow is designed to process examination results exported from DARA/LECTODARA optical mark recognition systems.

The input consists of a Microsoft Excel workbook (.xlsx) containing one record per examination and a set of descriptive fields organized by columns.

The scripts currently use the following fields (Minimum required fields):

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

---

# Additional Fields

Additional fields may be included in the source Excel file without affecting the workflow.

Examples:

- DNI
- NOMBRE
- APELLIDOS
- CONVOCATORIA
- TRIBUNAL
- SEDE
- MUNICIPIO
- OBSERVACIONES

These fields are ignored unless explicitly incorporated into a script.

Example:

```matlab
if ismember('DNI',T.Properties.VariableNames)
    dni = string(T.DNI);
end

## How to validate additional fields
In case these fields are required, it is advisable to include this validation routine in the scripts

Example:

if ismember('DNI',T.Properties.VariableNames)

    dni = string(T.DNI);

end

if ismember('NOMBRE',T.Properties.VariableNames)

    nombre = string(T.NOMBRE);

end

## Typical workflow

1. Split the master Excel file by subject.
2. Analyse centres within each subject.
3. Generate master statistical tables.
4. Generate global institutional analyses.
5. Generate subject-level reports.
6. Generate the final integrated institutional report.

## Workflow SCRIPST USED

1. `analisis_materia_a_materia.m`
2. `analisis_materia_v4Ab11.m`
3. `analisis_materia_centro_ordinaria3.m`
4. `analisis_pauV2.m`
5. `analisis_materia_v4Aa_PARAEXTRAORDINARIA.m`
6. `analisis_materia_v4_PARAEXTRAORDINARIA.m`
7. `INFORME_FINAL_NUEVOV2.m`



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
