# Sales data for the NextWork Excel + Claude projects

Sample sales orders for a small online electronics store, January to July 2026. Made up for learning; no real customers.

| File | What it is | Used in |
|---|---|---|
| [messy-sales.xlsx](messy-sales.xlsx) | 200 rows with duplicates, mixed date formats, blank cells, inconsistent product names, and extra spaces | Clean and Analyze a Messy Spreadsheet with Claude |
| [clean-sales.xlsx](clean-sales.xlsx) | The same orders after cleaning (190 rows; the Sales column is a formula, Quantity × Unit Price) | Build a Sales Dashboard with Claude in Excel |
| [july-sales.xlsx](july-sales.xlsx) | One new month of orders in the clean layout | Build a Sales Dashboard with Claude in Excel (Secret Mission) |

Download a file: open it above and select **Download raw file**.

`make_data.py` rebuilds every file from a fixed seed (`pip install openpyxl`, then `python make_data.py`). `ANSWER-KEY.md` lists the planted problems and the numbers a correct clean-up should find.
