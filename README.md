
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
