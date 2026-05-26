# Student Attendance Analytics

A Python project that analyses student attendance records, calculates attendance probabilities, and visualises patterns through a multi-sheet Excel workbook.

---

## Daily Work Logs (May 1 - May 25, 2026)

- **May 1:** Initialized the GitHub repository and planned the overall structure of the Student Attendance Analytics project.
- **May 2:** Defined the 12 student profiles with base attendance probabilities and assigned tier classifications (Excellent, Good, Average, Poor).
- **May 3:** Built the academic calendar generator to produce all Mon–Fri school days between Sep 2024 and May 2025 (195 days total).
- **May 4:** Wrote the core probabilistic attendance engine using Bernoulli trials with a fixed random seed for reproducibility.
- **May 5:** Integrated seasonal adjustment factors — winter illness dip (−8%), spring drop (−4%), and May recovery (+3%) — into the simulation.
- **May 6:** Added day-of-week effect modifiers (Monday +2%, Friday −4%) to reflect realistic weekly attendance patterns.
- **May 7:** Ran the simulation end-to-end and verified the 2,340-record DataFrame output against expected distributions.
- **May 8:** Designed the Excel workbook architecture — defined the 7-sheet layout, colour palette, and shared styling helper functions.
- **May 9:** Built the raw Attendance Log sheet with 2,340 rows, alternating row banding, tier badges, and Present/Absent status cells.
- **May 10:** Documented the first phase of results and pushed the initial codebase to the repository.
- **May 11:** Built the Student Summary sheet — per-student totals, attendance probability column, and day-of-week breakdown columns.
- **May 12:** Applied a three-way colour-scale conditional format to the probability column and embedded a clustered bar chart on the summary sheet.
- **May 13:** Created the Monthly Trends sheet with a class-level month-by-month attendance table and an annotated line chart.
- **May 14:** Built the per-student monthly heatmap table using a five-shade green-to-red encoding (≥95% dark green → <65% red).
- **May 15:** Developed the Day-of-Week Patterns sheet — weekday rate table, per-student grid, and a yellow-palette bar chart.
- **May 16:** Built the Weekly Tracker sheet with a 38-week attendance vs. absence rate table and dual-line chart.
- **May 17:** Created the At-Risk Students sheet — flagged all students below 80%, calculated days-needed-to-recover, and assigned risk levels.
- **May 18:** Added the coral alert summary banner to the At-Risk sheet and verified all threshold logic against raw data.
- **May 19:** Assembled the Dashboard sheet with four KPI cards (overall rate, student count, school days, at-risk count) and a tier breakdown table.
- **May 20:** Cross-checked all sheet values against the source DataFrame to confirm no formula or aggregation errors existed.
- **May 21:** Applied cell formatting, background colours, and border styles consistently across all seven sheets for a professional finish.
- **May 22:** Fixed the hardcoded output path from `/mnt/user-data/outputs/` to a local relative path for Windows compatibility.
- **May 23:** Tested the full script end-to-end on Windows (Python 3.14) and resolved a marker colour attribute error in the line chart.
- **May 24:** Conducted final proofreading of all sheet logic, chart labels, and conditional formatting rules across the workbook.
- **May 25:** Finalized README documentation, verified repository access links, and prepared for final submission.

Submitted By: Syed Suzain

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
