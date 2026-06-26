# Production Planning App

A 3-day automated production planning tool built with Python and Streamlit.

## Features
- Allocates production orders across 2 lines (Line 1: Roof + Wall, Line 2: Roof only)
- Prioritizes by SO ageing (oldest first)
- Minimizes thickness changeovers
- Targets configurable line utilisation %
- Outputs day-wise schedule to Excel

## How to Run
pip install streamlit pandas openpyxl
streamlit run app.py

## Input
Upload the `PLANNING_INPUT.xlsx` file with DEMAND and CAPACITY sheets.
