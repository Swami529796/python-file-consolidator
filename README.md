# Excel File Consolidator (Python)

Combine dozens of Excel and CSV files into one master workbook with a single command, with header validation and a full audit log. This is the Python rebuild of my [VBA File Consolidator](https://github.com/Swami529796/excel-file-consolidator).

![Python](https://img.shields.io/badge/Python-3776AB?style=flat&logo=python&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat&logo=pandas&logoColor=white)
![Tests](https://img.shields.io/badge/tests-6%20passing-brightgreen)
![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)

---

## The Problem

Every month, teams receive the same report from multiple regions, branches or vendors, each in a separate file. Someone opens every file, copies the data and pastes it into one master sheet. That is slow, error-prone and leaves no record of which row came from which file.

## The Solution

```bash
python consolidate.py sample-data
```

The script:

1. Reads every `.xlsx`, `.xlsm`, `.xls` and `.csv` file in the folder
2. Checks that each file's headers match the first file's
3. Appends all rows to a **Master** sheet, tagged with the source file and sheet name
4. Skips files that don't match instead of corrupting the output, and explains why on a **Log** sheet

## Sample Run

The `sample-data` folder holds four regional sales files. The West file deliberately uses `Qty` instead of `Units` so you can see validation in action.

```
Consolidation complete.
  Files found:      4
  Sheets loaded:    3
  Skipped / errors: 1
  Rows added:       120
  Time taken:       0.3 sec
  Output:           Master.xlsx
  ! Sales_West.xlsx: Header mismatch - column 4: expected 'Units', found 'Qty'
```

| File | Sheet | Rows Added | Status | Message |
|---|---|---|---|---|
| Sales_East.csv | Sales_East | 45 | Loaded | |
| Sales_North.xlsx | Sales | 40 | Loaded | |
| Sales_South.xlsx | Sales | 35 | Loaded | |
| Sales_West.xlsx | Sales | 0 | Skipped | Header mismatch - column 4: expected 'Units', found 'Qty' |

## VBA vs Python

Both versions solve the same problem with the same sample data and give the same 120 rows.

| | [VBA version](https://github.com/Swami529796/excel-file-consolidator) | Python version |
|---|---|---|
| Runs from | A button inside Excel | The command line, or a scheduled task |
| Needs installed | Excel only | Python, pandas, openpyxl |
| Time on the sample data | 8.6 sec | 0.3 sec |
| Lines of code | about 340 | about 220 |
| Opens each file in Excel | Yes | No, files are read directly |
| Automated tests | No | Yes, 6 tests with pytest |
| Best for | Colleagues who live in Excel and can't install software | Large volumes, scheduled runs, pipelines |

**My take:** VBA wins when the user must stay inside Excel. Python wins on speed, testing and anything that has to run unattended.

## How to Use

```bash
pip install -r requirements.txt
```

```bash
python consolidate.py "C:/Reports/June" --output June_Master.xlsx
```

| Option | Default | Description |
|---|---|---|
| `folder` | (required) | Folder that contains the files to merge |
| `-o`, `--output` | `Master.xlsx` | Output workbook |
| `--all-sheets` | off | Read every sheet of each workbook, not just the first |

Run the tests:

```bash
python -m pytest
```

## Assumptions

- Headers are in **row 1** and data starts in **row 2**
- The **first valid file** (in name order) defines the expected column layout
- Header matching ignores case and leading or trailing spaces

## Project Structure

```
python-file-consolidator/
├── consolidate.py            # the script
├── tests/
│   └── test_consolidate.py   # pytest tests
├── sample-data/              # demo input files (same as the VBA project)
├── requirements.txt
├── README.md
└── LICENSE
```

## Roadmap

- [ ] Option to include files from subfolders
- [ ] Column mapping, so `Qty` can be matched to `Units`
- [ ] Write to a database table as well as Excel

---

**Author:** [Swaminathan Narayanan](https://www.linkedin.com/in/swaminathan-narayanan-9781a7154/): Data Analytics | Excel & VBA Automation | Python | SQL | Power BI
