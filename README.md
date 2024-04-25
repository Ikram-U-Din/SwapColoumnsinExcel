# Excel Column Swap
This Python script swaps the positions of columns in an Excel file. It performs both horizontal and vertical flipping of data.

## Table of Contents

- [Introduction](#introduction)
- [Features](#features)
- [Requirements](#requirements)
- [Usage](#usage)
- [Example](#example)
- [Contributing](#contributing)

## Introduction

This script is designed to automate the process of swapping the positions of columns in an Excel file. It uses the `openpyxl` library to read and modify Excel files. The script can swap columns horizontally and then flip the data vertically.

## Features

- Swap the positions of specified columns in an Excel file.
- Flip the data vertically after swapping columns.

## Requirements

- Python 3.x
- `openpyxl` library

## Usage

1. Ensure that you have Python 3.x installed on your system.
2. Install the `openpyxl` library using pip:

   ```
   pip install openpyxl
   ```

3. Clone or download the repository to your local machine.
4. Run the script by providing the path to your Excel file as an argument.

   ```
   python excel_column_swap.py your_excel_file.xlsx
   ```

## Example

Suppose you have an Excel file named `data.xlsx` with the following columns:

| Col1 | Col2 | Col3 | Col4 | Col5 | Col6 | Col7 | Col8 |
|------|------|------|------|------|------|------|------|
|   A  |   B  |   C  |   D  |   E  |   F  |   G  |   H  |

Running the script with this file will swap the columns as follows:

| Col8 | Col7 | Col6 | Col5 | Col4 | Col3 | Col2 | Col1 |
|------|------|------|------|------|------|------|------|
|   H  |   G  |   F  |   E  |   D  |   C  |   B  |   A  |

The data will also be flipped vertically, resulting in the following:

| Col8 | Col7 | Col6 | Col5 | Col4 | Col3 | Col2 | Col1 |
|------|------|------|------|------|------|------|------|
|   A  |   B  |   C  |   D  |   E  |   F  |   G  |   H  |

## Contributing

Contributions are welcome! If you find any bugs or have suggestions for improvement, please open an issue or submit a pull request.
