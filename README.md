# Safety Quiz Checker

Scripts written in Python meant to automate the checking of responses provided in an Excel workbook extracted from FormSG.

## Setup Guide

### Step 1: Install Python

The version used to build this was Python 3.9.13. It is recommended to download this version of Python to ensure compability.

Python can be installed [here.](https://www.python.org/downloads/), and more information about Python 3.9.13 can be found [here.](https://www.python.org/downloads/release/python-3913/)

**Note: Remember to select the "Add Python3.9 to PATH" checkbox. This saves you a lot of time.**

### Step 2: Clone the repository

It is highly recommended to install python dependencies in a virtual environment.

### Step 3: Install required dependencies

```console
pip install -r requirements.txt
```

## Usage Guide

### Preparation Work

- You are required to first submit a response that has all the correct answers, before you allow others to submit responses. This is because the responses are graded with reference to the first response. The `Company` and `Name` of said response does not matter.

- When extracting the form into excel, save your file in `.xlsx` format, and place it within the `data/` folder. `.csv` files are fine as well, but not recommended.

- If there were any files that were present in `data/`, clear them before inserting the new excel file into the folder.

### Results

The original excel file will have a new sheet titled `results`, which will highlight wrong answers and the score attained by each response. This excel file will be accessible in the `results/` folder.

The file will be in `.xlsx` format for highlighting of cells to be possible.
