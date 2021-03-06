# Eagle BOM Converter
a tool to convert CSV output file of the BOM.ulp into neatly arranged excel file with an option to add distributor price using Octopart excel add-in.

## Features
- convert a CSV BOM (exported from eagle) to excel format.
- Removes unnecessary attributes.
- Adding an index column.
- Adding an header row. 
- Removes rows with EXCLUDE attribute from BOM.
- Generate an excel BOM with Octopart formulas and price estimate. 

## Dependencies
* **Python3.**
* **pandas.**  
Install using pip:   
```shell
pip install pandas
```
* **openpyxl.**   
Install using pip:
```shell
pip install openpyxl
```
* **Octopart add-in.**  
In order to install the add-in follow the instructions in the link: https://octopart.com/excel

## Usage

1. add Part Attributes in eagle library.  
This script is searching for specific attributes in the csv file:   
    - Parts = part designator.     
    - Qty = part quantity.  
    - MF  = part manufacturer.   
    - MPN = manufacturer part number. 
    - VALUE = I use it for capcitors and resistors values.
    - FOOTPRINT =  part package.
    - DESCRIPTION = I use the Description from Digi-key.

You can add those attributes to your parts or change the line of code that contains the attributes:
```python
newDf = bomDf[['Qty', 'Parts','MF','MPN','VALUE','FOOTPRINT','DESCRIPTION']].copy()
```
2. Export CSV BOM from Eagle:  
To generate a BOM file from Eagle make sure you are in the Schematic Editor and go to **File > Export > BOM.** This will open up a new window where you can configure how you want the BOM file to look like. Check Values and CSV option. Next you can save the bom.csv file to the desired location by clicking the Save button.

![image](docs/pic/bomExport.JPG)

3. Run the python script from command line:
```shell
[repo_path]/python bomToolGuiV3.py  
```
you will get this window:  

![image](docs/pic/app.JPG)

Press the “select a file” button and chose the file to convert, the path for the file will appear in the text box. to generate the BOM press the “RUN” button on completion the text in path text box will change to “done select new file” the new file is saved in the same folder as the original CSV file.  
**Octopart** - To generate a BOM with OCTOPART_DISTRIBUTOR_PRICE formula and total price caluclation column check the “Octopart” check box.
## TO DO
- [ ] improve column with auto size fit.
- [ ] easy attribute customization.