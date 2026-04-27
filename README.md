# NHS Hospital Admissions Dashboard 2012 to 2024

## COMP4037 Research Methods Coursework 2

Student Name: Yash Varma  
Student ID: 20794563  
University: University of Nottingham  
Module: COMP4037 Research Methods  
Coursework: Coursework 2, Data Visualization  

## Project Overview

This project analyses NHS Hospital Admissions data to understand how hospital admissions changed before, during, and after the COVID-19 lockdown period.

The main output is an interactive Streamlit dashboard containing six visualization types:

1. Hierarchical treemap
2. Pre COVID baseline heatmap
3. Radar chart
4. Violin plot
5. Parallel coordinates chart
6. Supporting stacked admission route chart

## Research Question

Which diagnosis categories were most disrupted by the 2020 to 2021 COVID-19 lockdown, and what unexpected admission or emergency-care anomalies emerged?

## Key Findings

The dashboard highlights three main findings:

1. F50 to F59, Eating and Behavioural Syndromes, increased during the lockdown period while many other categories declined.
2. M00 to M25, Arthropathies, showed an emergency-care paradox: total admissions fell, but emergency share increased.
3. H25 to H28, Cataracts and Lens Disorders, fell sharply during lockdown but rebounded strongly by 2023 to 2024, suggesting planned-care backlog and recovery pressure.

Overall, the lockdown period redistributed hospital admissions across clinical categories rather than causing a simple uniform increase or decrease.

## Tools and Libraries

This dashboard was built using:

- Python
- Streamlit
- Pandas
- NumPy
- Plotly
- OpenPyXL

## Project Structure

Your project folder should look like this:

```text
nhs-hospital-admissions-dashboard
│
├── app_dashboard.py
├── requirements.txt
├── README.md
│
├── NHS Hospital Admissions
│   ├── hosp-epis-stat-admi-diag-2012-13-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2013-14-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2014-15-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2015-16-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2016-17-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2017-18-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2018-19-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2019-20-tab supp.xlsx
│   ├── hosp-epis-stat-admi-diag-2020-21-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2021-22-tab.xlsx
│   ├── hosp-epis-stat-admi-diag-2022-23-tab_V2.xlsx
│   └── hosp-epis-stat-admi-diag-2023-24-tab.xlsx
│
└── .streamlit
    └── config.toml
```

## Required Files

The repository should contain:

```text
app_dashboard.py
requirements.txt
README.md
NHS Hospital Admissions/
.streamlit/config.toml
```

## Requirements File

Create a file named exactly:

```text
requirements.txt
```

Paste this inside:

```txt
streamlit==1.39.0
pandas==2.2.3
numpy==2.1.2
plotly==5.24.1
openpyxl==3.1.5
```

If deployment gives dependency issues, use this simpler version instead:

```txt
streamlit
pandas
numpy
plotly
openpyxl
```

## Streamlit Theme Configuration

Create a folder named:

```text
.streamlit
```

Inside it, create a file named:

```text
config.toml
```

Paste this inside:

```toml
[theme]
base = "light"
backgroundColor = "#FFFFFF"
secondaryBackgroundColor = "#F7F7F7"
textColor = "#1F1F1F"
primaryColor = "#355CDE"
font = "sans serif"
```

This keeps the dashboard white and clean.

## Data Folder

The dashboard expects the data folder to be named:

```text
NHS Hospital Admissions
```

This folder should be placed in the same directory as:

```text
app_dashboard.py
```

The dashboard code automatically looks for this folder:

```python
local_data_path = Path(__file__).resolve().parent / "NHS Hospital Admissions"
default_path = str(local_data_path if local_data_path.exists() else Path.cwd())
```

This works both locally and on Streamlit Cloud, as long as the folder is included in the project.

## How to Run Locally

### Step 1: Open Terminal in the Project Folder

Open the folder containing:

```text
app_dashboard.py
requirements.txt
NHS Hospital Admissions/
```

On Windows, open the folder in File Explorer, click the address bar, type:

```text
cmd
```

Then press Enter.

### Step 2: Install Dependencies

Run:

```bash
pip install -r requirements.txt
```

If `pip` is not recognised, run:

```bash
python -m pip install -r requirements.txt
```

### Step 3: Run the Dashboard

Run:

```bash
streamlit run app_dashboard.py
```

If `streamlit` is not recognised, run:

```bash
python -m streamlit run app_dashboard.py
```

### Step 4: Open the Dashboard

The dashboard should open automatically in your browser.

If it does not open automatically, go to:

```text
http://localhost:8501
```

## How to Use the Dashboard

The sidebar contains the main controls:

### Years

Select the financial years to include in the dashboard.

For the main COVID lockdown analysis, include:

```text
2019-20
2020-21
2021-22
2023-24
```

### Age Groups

Select age groups for age-based analysis.

For older patient burden, select:

```text
65 to 69
70 to 74
75 to 79
80 to 84
85 to 89
90+
```

### Abstraction Level

The dashboard supports:

```text
Summary diagnosis
3-character diagnosis
4-character diagnosis
```

The main report findings use the Primary Diagnosis Summary level because the key anomalies are clear and readable at this level.

### Heatmap Metric

The heatmap supports two metrics:

```text
Admissions % change from 2019-20 baseline
Emergency share percentage point change from 2019-20 baseline
```

Use the first option to show total admission disruption.

Use the second option to show emergency-care pressure.

## Main Visualizations

### Treemap

The treemap shows the ICD-10 hierarchy.

- Size represents admission burden.
- Colour represents emergency admission share.
- It helps reveal whether pressure is concentrated in broad or specific diagnosis categories.

### Baseline Heatmap

The heatmap compares each diagnosis category against the 2019 to 2020 pre COVID baseline.

- Red means admissions decreased.
- Blue means admissions increased.
- White means little change.
- Grey means unavailable category-year data.
- Amber highlights the 2020 to 2021 lockdown year.

### Radar Chart

The radar chart compares selected categories across multiple pressure indicators:

- Selected age burden
- Emergency share
- Mean waiting time
- Mean length of stay
- Mean age

### Violin Plot

The violin plot shows how admissions are distributed across age groups.

It helps identify whether particular age groups dominate certain diagnosis categories.

### Parallel Coordinates

The parallel coordinates chart supports multivariable comparison across:

- Finished Consultant Episodes
- Finished Admission Episodes
- Emergency admissions
- Mean waiting time
- Mean length of stay
- Mean age
- Selected age-group burden

### Stacked Admission Route Chart

The stacked chart shows emergency, waiting list, planned, and other admission routes.

This is used as a supporting comparison.

## Suggested Video Research Question

Which diagnosis categories were most disrupted by the 2020 to 2021 COVID-19 lockdown, and what unexpected admission or emergency-care anomalies emerged?

## Suggested Video Script

This dashboard visualises NHS hospital admissions across six chart types, covering financial years 2012 to 2024. My research question is: which diagnosis categories were most disrupted by the 2020 to 2021 COVID-19 lockdown, and what unexpected admission or emergency-care anomalies emerged?

The sidebar lets me filter by year, age group, abstraction level, and metric. The treemap at the top shows the hierarchical burden of each diagnosis category, where size represents admission volume and colour represents emergency admission share. Below it, the baseline heatmap shows percentage change against the 2019 to 2020 pre-pandemic baseline.

In the heatmap, red means admissions fell, blue means they rose, and white means little or no change. Grey cells show unavailable category-year data. The amber column highlights 2020 to 2021, the main lockdown disruption year.

The first anomaly appears in the F-chapter rows. Most categories are red in 2020 to 2021, but F50 to F59, Eating and Behavioural Syndromes, rises by 15 percent, from 4,516 to 5,200 admissions. By 2021 to 2022, it reaches 39 percent above baseline, making it one of the few non-COVID categories to increase during the lockdown period.

The second anomaly is M00 to M25, Arthropathies. Admissions fell by 53 percent, but emergency share rose from 13 to 23 percent, showing that urgent musculoskeletal cases formed a larger share of the remaining admissions.

Finally, H25 to H28, Cataracts and Lens Disorders, fell by 46 percent during 2020 to 2021, but rebounded to around 55 percent above baseline by 2023 to 2024. This suggests a planned-care backlog and later recovery pressure.

Overall, the dashboard shows that lockdown did not reduce admissions uniformly. It redistributed hospital pressure across categories in different ways.

## Deployment to Streamlit Community Cloud

### Step 1: Upload Project to GitHub

Make sure your GitHub repository contains:

```text
app_dashboard.py
requirements.txt
README.md
NHS Hospital Admissions/
.streamlit/config.toml
```

### Step 2: Go to Streamlit Community Cloud

Open Streamlit Community Cloud and create a new app.

### Step 3: Select Repository

Choose your GitHub repository.

### Step 4: Set Main File

Set the main file path to:

```text
app_dashboard.py
```

### Step 5: Deploy

Click deploy.

The first load may take some time because the dashboard reads several Excel files.

## GitHub Terminal Commands

Run these inside the project folder:

```bash
git init
git add .
git commit -m "Add NHS hospital admissions dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/nhs-hospital-admissions-dashboard.git
git push -u origin main
```

Replace:

```text
YOUR_USERNAME
```

with your real GitHub username.

## Troubleshooting

### Streamlit is not recognised

Run:

```bash
python -m streamlit run app_dashboard.py
```

### Pip is not recognised

Run:

```bash
python -m pip install -r requirements.txt
```

### No Excel files found

Check that the folder is named exactly:

```text
NHS Hospital Admissions
```

and that it is in the same folder as:

```text
app_dashboard.py
```

### Dashboard loads slowly

The first load may be slow because the app reads multiple Excel files. Streamlit caching should make later loads faster.

### GitHub rejects a file

GitHub blocks files larger than 100 MB. If this happens, remove the large file or use a smaller processed dataset.

## Coursework Note

This dashboard was created for COMP4037 Research Methods Coursework 2 at the University of Nottingham.
