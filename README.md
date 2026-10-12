# Project-1--QUIZ-Exam-Improvement

**JOSHUA KURT M. MANZANO**  
**BSIT 3-3**

## Overview

Project 1 is a Python-based Quiz/Exam Improvement tool with a graphical user interface (GUI). It analyzes quiz or exam responses and flags patterns that may need review for possible cheating or duplicate submissions.

The program supports **CSV (`.csv`)** and **Excel (`.xlsx`, `.xlsm`)** files.

## Requirements

- Python 3.9 or later recommended
- `pandas` (if used by your project environment)
- `openpyxl` — required for reading Excel files
- `tkinter` — used for the graphical interface; it is included with many standard Python installations

Install the Excel dependency with:

```bash
python -m pip install openpyxl
```

> Note: The current `proj1.py` implementation uses Python's built-in `csv` module for CSV files and `openpyxl` for Excel files. It does not directly import pandas, so pandas is not required by this script as currently written.

## How to Run

1. Make sure Python is installed.
2. Install the Excel dependency if you plan to open `.xlsx` or `.xlsm` files.
3. Open a terminal or command prompt in the project folder.
4. Run the program:

   ```bash
   python proj1.py
   ```

5. The GUI window will open. Click **Check for Cheating** and select the responses file.
6. If you have an answer key, choose **Yes** when prompted and select the answer-key file. Providing a key makes the analysis more reliable.
7. Review the report shown in the application. A text report is also saved automatically.

## Input Files

### Responses file

Use a CSV or Excel spreadsheet with:
- One header row containing the column names
- One row per student response
- Question columns containing the students' answers
- Optional identity columns such as `Name`, `Student Name`, `Email`, or `Student ID`
- An optional `Timestamp` column if submission timing is available

The program treats recognized identity columns and the `Timestamp` and `Score` columns as metadata rather than question answers.

### Answer key (optional)

The answer-key file should contain:
- One header row with question names matching the response file
- One row containing the correct answer for each question

The answer key must use the same question-column headers as the responses file for the answers to match correctly.

If no answer key is supplied, the program treats the most common non-empty answer for each question as the correct answer. This is less reliable and may affect the results.

## What the Program Checks

- **Matching wrong answers:** Flags pairs of students who share enough identical incorrect answers.
- **Similar written answers:** Flags sufficiently similar long-form answers.
- **Close submission times:** Adds supporting points when two flagged students submitted within the configured time window.
- **Repeated submissions:** Flags multiple submissions associated with the same recognized identity.

The thresholds can be adjusted in the **EDIT ONLY THIS SECTION** near the top of `proj1.py`:

- `MIN_SHARED_WRONG`
- `SHARED_WRONG_RATIO`
- `TEXT_SIMILARITY`
- `MIN_TEXT_LENGTH`
- `CLOSE_SUBMIT_SECONDS`

## Output

Reports are saved automatically in:

```text
Documents/Quiz_Exam_Project/
```

The output file name follows this format:

```text
Cheating_Report_YYYYMMDD_HHMM.txt
```

The report includes the number of responses checked, notes about the analysis, flagged patterns, and reasons for each flag.

## Command-Line Usage (Optional)

You can run the checker without opening the GUI by supplying the responses file:

```bash
python proj1.py check path/to/responses.csv
```

To include an answer key:

```bash
python proj1.py check path/to/responses.xlsx path/to/answer_key.xlsx
```

Use quoted paths if a file or folder name contains spaces.

## Supported Formats and Limitations

- Supported: `.csv`, `.xlsx`, `.xlsm`
- Not supported: legacy Excel `.xls` files. Open the file in Excel and use **Save As** to convert it to `.xlsx` or `.csv`.
- Excel workbooks are read from their first worksheet.
- At least two student responses are needed for a pairwise comparison.

## Important Notice

This tool identifies patterns that may warrant further review. **A flag is not proof of cheating.** Similar answers can occur for legitimate reasons. Review the original submissions and other relevant context before drawing conclusions.

## Project File

- `proj1.py` — main Python script containing the GUI and response-analysis logic.
