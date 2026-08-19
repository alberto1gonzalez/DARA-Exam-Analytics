
# DARA-Exam-Analytics

MATLAB toolkit for analysing examination results obtained from either DARA/LECTODARA optical mark recognition systems or optical readers, generating statistics, rankings, reports, maps and GIS products.

The workflow was originally developed for the Spanish University Entrance Examination (PAU) but can be adapted to any assessment process based on optical mark recognition (OMR).

## Overview

This repository contains a collection of MATLAB scripts developed for the statistical analysis of examination results exported from DARA/LECTODARA optical readers. The workflow was originally developed for the Spanish PAU (University Entrance Examination) but can be adapted to any assessment process based on optical mark recognition (OMR).

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

# Required MATLAB Toolboxes

Mandatory:

- MATLAB
- MATLAB Report Generator

Optional:

- Statistics and Machine Learning Toolbox
- Mapping Toolbox

# Input Data Structure

The input consists of a Microsoft Excel workbook (.xlsx) containing one record per examination.

Minimum required fields:

| Field | Description |
|---------|---------|
| CENTRO | Educational centre identifier |
| TIPO | Subject code |
| MATERIA | Subject name |
| TIENEEXAMEN | Examination taken (SI / NO) |
| CALIFICACION | Score (0–10) |
| COORDENADA_X | X Coordinate |
| COORDENADA_Y | Y Coordinate |

Example:

CENTRO | TIPO | MATERIA | TIENEEXAMEN | CALIFICACION
---------|---------|---------|---------|---------
CENTRO_A | BIO | Biología | SI | 6.75
CENTRO_B | BIO | Biología | SI | 7.80
CENTRO_C | FIS | Física | SI | 5.25
CENTRO_D | FIS | Física | NO | 

The workflow automatically filters:
- missing scores,
- examinations not taken,
- invalid scores outside the range 0–10.

Additional fields may be present and are preserved, although they are not required by the current workflow.

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

How to Validate Additional Fields

If additional fields are required, it is advisable to include validation routines similar to the following:

Example:
if ismember('DNI',T.Properties.VariableNames)
    dni = string(T.DNI);
end

if ismember('NOMBRE',T.Properties.VariableNames)
    nombre = string(T.NOMBRE);
end

if ismember('APELLIDOS',T.Properties.VariableNames)
    apellidos = string(T.APELLIDOS);
end
``
# Example Dataset
examples/
    ejemplo_datos_1000_registros.xlsx

# Workflow Scripts Used

1. `analisis_materia_a_materia.m`
2. `analisis_materia_v4Ab11.m`
3. `analisis_materia_centro_ordinaria3.m`
4. `analisis_pauV2.m`
5. `analisis_materia_v4Aa_PARAEXTRAORDINARIA.m`
6. `analisis_materia_v4_PARAEXTRAORDINARIA.m`
7. `INFORME_FINAL_NUEVOV2.m`

# Workflow Outputs


# Typical Workflow

1. Split the master Excel file by subject.
2. Analyse centres within each subject.
3. Generate master statistical tables.
4. Generate global institutional analyses.
5. Generate subject-level reports.
6. Generate the final integrated institutional report.

# Workflow Scripts Used

      ## Phase 1. Split Master Excel File
    
          Script:
          ```text
          analisis_materia_a_materia.m
          ```
          
          Generates:
          
          ```text
          MATERIAS\
              BIO\
              FIS\
              MAT\
              ...
          ```
    
    ## Phase 2. Subject and Centre Analysis
    
          Script: 
          ```text
          analisis_materia_v4Ab11.m
          ```
          
          Generates:    
          ```text
          RESULTADOS\
              INFORME_*.*ocx
              Resumen_Centros.xlsx*    hist_*.png
          ```
          
    ## Phase 3. Master Statistical Table
    
          Script:       
          ```text
          analisis_materia_centro_ordinaria3.m
          ```
          
          Generates:
          
          ```text
          RESULTADOS_MAESTROS\
              resumenes_excel\
                  resumen_materia_centro.xlsx
          ```
    
            Contains:
            
            - Mean
            - Median
            - Standard deviation
            - Coefficient of variation
            - Composite index
            - Global ranking
            - Subject ranking
    
    ## Phase 4. Institutional Analysis
    
          Script:      
          ```text
          analisis_pauV2.m
          ```
          
          Generates:
          
          ```text
          RESULTADOS_PAU\
              ranking_centros.xlsx
              histograma_global.png
              mapa_centros.png
              resultados_centros.shp
              informe_pauV2.docx
          ```
    
    ## Phase 5. Global Subject Analysis
    
          Script:        
          ```text
          analisis_materia_v4Aa_PARAEXTRAORDINARIA.m
          ```
          
          Generates:
          
          ```text
          resultados_extraordinaria\
              INFORME_GLOBAL.docx
              hist_*.png
          ```
    
    *# Phase 6. Subject Reports with EM* Graphics
    
          Script:   
          ```text
          analis*s_materia*v4_PARA*XTRAORDINARIA.m
          ```
          
          Gener*tes:
          
          ```text
          resultados_extraordi*aria\
              hist_*.png
              hist_*.*mf
              informe_*.docx
              INFORME_GLOBAL.docx
          ```
    
    ## Phase 7. Integrated Institutional Report
    
          Script:          
          ```text
          INFORME_FINAL_NUEVOV2.m
          ```
          
          Generates:
          
          ```text
          INFORME_GLOBAL_EXTRAORDINARIA_V2.docx
          ```

# Workflow Outputs

The toolkit generates:

- Subject-level reports (.docx)
- Global reports (.docx)
- Histograms (.png)
- Vector histograms (.emf)
- Statistical summaries (.xlsx)
- Centre rankings (.xlsx)
- GIS products (.shp)
- Final institutional reports
``

## License

MIT License


