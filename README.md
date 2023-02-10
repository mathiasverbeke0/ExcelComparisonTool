# Excel Comparison Tool
This application is used to compare the contents of two csv files and identify any common lines. The comparison is performed line by line, and the result is saved in a new excel file. The original excel files are first converted to csv files, and the comparison is performed on the csv files.

## Requirements
* Python 3.x
* Openpyxl
* Pandas

## Installation
1. Clone the repository
```bash
git clone https://github.com/mathiasverbeke0/ExcelComparisonTool.git
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
python <script_name>.py <filename1> <filename2>
```
Where <script_name>.py is the name of the script and <filename1> and <filename2> are the names of the two excel files you want to compare.

## Limitations
* The script currently only supports Excel files in the .xlsx format.
* The csv files are created in the same directory as the script, so make sure to have the necessary write persmission in that directory.
* The script only compares the contents of the csv files data line by data line. Any other comparison methods (e.g. column-wise comparison) are not currently supported.

## Contributing
1. Fork the repository
2. Create your feature branch (git checkout -b my-new-feature)
3. Commit your changes (git commit -am 'Add some feature')
4. Push to the branch (git push origin my-new-feature)
5. Create a new Pull Request