# Final Performance Assessment (SDC205 5.5)

**Student:** Donavan McFadden | **Student ID:** Donmcf8357

`finalExam.py` lives in `C:\FinalExam` next to `final.csv` and runs four functions:

| Function | What it does |
|---|---|
| `askUser()` | A loop asks for five numbers, adds them to a running total, and displays the total. |
| `askIncome()` | A loop asks for the names and annual incomes of five people and appends each one to `final.csv` on a new line. The file goes from 4 entries to 9. |
| `excelPie()` | Uses openpyxl to copy the data into a new `final.xlsx` and add a pie chart titled with my Student ID and today's date. |
| `verticalBar()` | Uses matplotlib to show a vertical bar graph of the same data with the same title. |

## Files
- `finalExam.py` – the program
- `final.csv` – the data file after the run (9 entries)
- `final.xlsx` – the Excel workbook with the data and pie chart, made by the program
- `screenshots/` – IDLE output, Excel pie chart, and the vertical bar graph

## How to run
```
cd C:\FinalExam
python finalExam.py
```
Requires `openpyxl` and `matplotlib` (`pip install openpyxl matplotlib`).
