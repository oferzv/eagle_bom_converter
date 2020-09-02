# eagle_bom_converter
a simple script to convert eagle BOM ULP csv output file in to excel format and add distributor price using Octopart excel add-in

## Features
- convert a CSV BOM (exported fro eagle) to excel.
- add an index colume. 
- Removes rows with EXCLUDE attribute from BOM.
- Generate a excel BOM with Octopart formulas and price estimate. 

## Dependencies
### Install openpyxl using pip

```shell
pip install openpyxl
```

### Octopart add-in 

## Usage

1. add Part Attributes in eagle library:
This script is searching for specific attributes in the csv file: 
    Qty, Parts, MF, MPN, VALUE, FOOTPRINT, DESCRIPTION
You can add those attributes to your parts or change the line of code that contains the attributes:
```python
  newDf = bomDf[['Qty', 'Parts','MF','MPN','VALUE','FOOTPRINT','DESCRIPTION']].copy()
```

## TO DO
- Improve column with auto size fit