# Excel Comparison Tool
This application is used to compare the contents of two csv files and identify any common or unique lines. The comparison is performed line by line, and the result is saved in a new excel file. The original excel files are first converted to csv files, and the comparison is performed on the csv files.

## Requirements
* Python 3.x
* Openpyxl
* Pandas

## Installation
1. Clone the repository
```bash
git clone git@github.com:mathiasverbeke0/ExcelComparisonTool.git
```

2. Install required packages
```bash
pip install openpyxl
pip install pandas
pip install numpy
```

## Usage
The script takes in two file names as command line arguments. To run the script, use the following command in your terminal:

```bash
usage: python ExcelComparisonTool.py [-h] [-u] filename1 filename2

positional arguments:
  filename1         name of the first Excel file
  filename2         name of the second Excel file

options:
  -h, --help        show this help message and exit
  -u, --unique      use this option to see what lines in file1 are not present in file2
```

## Examples
```bash
python ExcelComparisonTool.py file1.xlsx file2.xlsx
```
This command will compare the contents of 'file1.xlsx' and 'file2.xlsx' and identify the common lines. 

```bash
python ExcelComparisonTool.py -u file1.xlsx file2.xlsx
```
This command will compare the contents of 'file1.xlsx' and 'file2.xlsx' and identify the lines in 'file1.xlsx' that are not present in 'file2.xlsx'.

## Limitations
* The script currently only supports Excel files in the .xlsx format.
* The csv files are created in the same directory as the .xlsx files, so make sure to have the necessary write persmission in that directory.
* The script only compares the contents of the csv files data line by data line. Any other comparison methods (e.g. column-wise comparison) are not currently supported.

## Contributing
1. Fork the repository
2. Create your feature branch (git checkout -b my-new-feature)
3. Commit your changes (git commit -am 'Add some feature')
4. Push to the branch (git push origin my-new-feature)
5. Create a new Pull Request