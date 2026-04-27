# NHS Hospital Admissions Dashboard

Interactive Streamlit dashboard for COMP4037 Research Methods Coursework 2.

Student: Yash Varma  
Student ID: 20794563  
University: University of Nottingham  

## Research Question

Which diagnosis categories were most disrupted by the 2020 to 2021 COVID-19 lockdown, and what unexpected admission or emergency-care anomalies emerged?

## Overview

This dashboard analyses NHS hospital admissions data from 2012 to 2024. It uses the 2019 to 2020 financial year as the pre-pandemic baseline and compares admission changes during and after the COVID-19 lockdown period.

## Visualisations

The dashboard includes:

- Hierarchical treemap
- Pre-COVID baseline heatmap
- Radar chart
- Violin plot
- Parallel coordinates chart
- Admission route composition chart

## Key Findings

- F50 to F59, Eating and Behavioural Syndromes, increased during the lockdown period.
- M00 to M25, Arthropathies, showed an emergency-care paradox: total admissions fell, but emergency share increased.
- H25 to H28, Cataracts and Lens Disorders, fell sharply during lockdown and later rebounded strongly.

## Project Structure

```text
project-folder
│
├── app_dashboard.py
├── requirements.txt
├── README.md
│
├── NHS Hospital Admissions
│   ├── hosp-epis-stat-admi-diag-2012-13-tab.xlsx
│   ├── ...
│   └── hosp-epis-stat-admi-diag-2023-24-tab.xlsx
│
└── .streamlit
    └── config.toml
```

## Installation

Install the required packages:

```bash
pip install -r requirements.txt
```

## Run Locally

Start the dashboard:

```bash
streamlit run app_dashboard.py
```

Then open:

```text
http://localhost:8501
```

## Requirements

```txt
streamlit
pandas
numpy
plotly
openpyxl
```

## Data Folder

The dashboard expects the data folder to be named:

```text
NHS Hospital Admissions
```

and placed in the same folder as `app_dashboard.py`.

## Deployment

This app can be deployed on Streamlit Community Cloud. Use:

```text
app_dashboard.py
```

as the main file.
