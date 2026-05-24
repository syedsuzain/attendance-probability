# Student Attendance Analytics

A Python project that analyses student attendance records, calculates attendance probabilities, and visualises patterns through a multi-sheet Excel workbook.

---

## What It Does

- Simulates a full academic year (Sep 2024 – May 2025) for 12 students across 195 school days
- Calculates the probability of each student attending on any given day
- Applies seasonal and day-of-week adjustments to model realistic patterns
- Outputs a 7-sheet Excel workbook with charts, heatmaps, and an at-risk report

---

## Project Timeline

### Day 1 — Setup & Data Engine
- Defined 12 student profiles with base attendance probabilities and tier classification
- Built the academic calendar (Mon–Fri only, 195 days)
- Wrote the probabilistic attendance simulator using Bernoulli trials with seasonal adjustments

### Day 2 — Excel Foundation & Raw Log
- Designed the 7-sheet workbook layout and colour palette
- Built shared helper functions for styling (fills, fonts, borders, alignment)
- Created the full raw **Attendance Log** sheet (2,340 rows, banded rows, status badges)

### Day 3 — Analytics Sheets
- **Student Summary** — per-student probability table with day-of-week columns and bar chart
- **Monthly Trends** — class-level line chart + per-student monthly heatmap
- **Day-of-Week Patterns** — weekday bias analysis with bar chart (Friday dip clearly visible)

### Day 4 — Tracker & Risk Report
- **Weekly Tracker** — dual-line chart (attendance vs absence) across 38 weeks
- **At-Risk Report** — flagged students below 80%, computed days-needed-to-recover, assigned risk levels
- Fixed the output path from `/mnt/user-data/outputs/` to a local relative path for Windows compatibility

### Day 5 — Dashboard & Polish
- Built the **Dashboard** sheet with KPI cards, tier breakdown table, and navigation hints
- Reviewed all sheets for consistent styling and formatting
- Tested end-to-end on Windows (Python 3.14) and confirmed clean output

---

## Output

Running the script produces `Student_Attendance_Analysis.xlsx` with these sheets:

| Sheet | Contents |
|---|---|
| 🏠 Dashboard | KPI cards, tier summary |
| 📊 Student Summary | Probabilities, bar chart |
| 📅 Monthly Trends | Line chart, heatmap |
| 📆 Day-of-Week | Weekday patterns, bar chart |
| 📈 Weekly Tracker | Dual-line chart across 38 weeks |
| 📋 Attendance Log | Full raw records |
| ⚠️ At-Risk Report | Students below 80% threshold |

---

## How to Run

```bash
pip install pandas openpyxl numpy
python attendance_analysis.py
```

The file saves in the same directory as the script.

---

## Tech Stack

| Tool | Use |
|---|---|
| Python 3.x | Core language |
| pandas | Aggregations and pivot tables |
| openpyxl | Excel workbook creation and charts |
| numpy / random | Probability simulation |
