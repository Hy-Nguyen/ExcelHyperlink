# Need to install openpyxl
`pip install openpyxl` or `pip3 install openpyxl`

# if that doesn't work try the following:

`python3 -m venv venv`

`source venv/bin/activate`

`pip install openpyxl`

`python script.py`


# Variables to be changed:
### sheet1 = wb["Lists"]
Change "Lists" to name of the sheet containg the lists
### sheet2 = wb["Links"]
Change "Lists to name of the sheet containing hyperlinks
### file_path = "/Users/hy/Code/DavidExcel/SampleData.xlsx"
Change to match system path of Excel Dock
