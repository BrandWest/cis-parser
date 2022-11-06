# cis-parser
This project parses the CIS Benchmarks into a CSV. 

# Requirements

PyPDF2==2.11.1
OpenPyxl==3.0.10

# To Use

1. Clone the repository
2. Install the requirements with ```pip -m install -r requirements.txt ```
3. Download the CIS Benchmarks files from https://www.cisecurity.org/cis-benchmarks/
4a. Using the CLI run ```python3 cis_parser.py -s <source file> -f <destination filename>```
  ALTERNATIVE
4b. Using the CLI to open a GUI run ```python3 cis_parser.py -g```


# Edge Cases
Note: Not all edge cases are accounted for and there could be some instances of failures or misinterpreted values. These can be remediated by opening an issues and provided the CIS Benchmark filename, specific benchmark, and output from the parser.
  
One issue that has been identified with PyPDF2 is that some PDF extractions gain spaces in words that are used to identify the specific content. This can be bypassed by resaving the PDF using an alternate means (e.g.: Print To File). An issues has been opened with PyPDF which can be found here: https://github.com/py-pdf/PyPDF2/issues/1424

# To Do
Implement logic for the Excel XLSX files.
