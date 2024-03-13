# Appraisal Results Generator

Scripts written in Python which will process raw appraisal results data.

## Setup Guide

### Step 1: Install Python

The version used to build this was Python 3.8.10. It is recommended to download this version of Python to ensure compatibility.

Python can be installed [here](https://www.python.org/downloads/), and more information about Python can be found here as well.

**Note: Remember to select the "Add Python3.8 to PATH" checkbox. This saves you a lot of time.**

### Step 2: Clone the repository

It is highly recommended to install Python dependencies in a virtual environment.

### Step 3: Install required dependencies

```console
pip install -r requirements.txt
```

Be sure to have `requirements.txt` in the current working directory for this command to work.

## Usage Guide

### Preparation Work

- Google Forms should have been used to collect results. Results extracted from *FormSG* follow a different format and is incompatible with these scripts.
- All peer appraisal results should be collected in the same workbook, and the same should be followed for all commander appraisal results.
  - Sheets should be named `PLT_1`, `PLT_2`, `PLT_3`, `PLT_4`, `PLT_5`
- `helpers.py` must be in the same directory as `peer-generation.py` and `comd_generation.py` in order for the codes to work.
- All workbooks must be in the same directory as all Python files in order for codes to work.

**Note:** the 2 `.ipynb` files are there for reference. They are my first drafts and have served me well for the 42/23 and 43/24 batches. I will be retiring
them in favour of the `.py` files.

### Steps

#### Peer Appraisal

1. Ensure all preparation work has been done
2. Change the input file name in `peer_generation.py` to the name of the excel file
3. Change the output file name in `peer_generation.py` to a name that you want. This will be the name of the excel file generated
4. Add in all 4-digit numbers of out-of-course personnel
5. Run the code
6. Check that all results have been reflected in the excel workbook

#### Commander Appraisal

1. Ensure all preparation work has been done
2. Change the input file name in `comd_generation.py` to the name of the excel file
3. Change the output file name in `comd_generation.py` to a name that you want. This will be the name of the excel file generated
4. Add in all 4-digit numbers of out-of-course personnel
5. Run the code
6. Check that all results have been reflected in the excel workbook

### Results

The output files will show up in the same directory as the codes.
