# Green-Stock Analysis

### BACKGROUND

> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

The analysis below will provide insight into the following:
- DevOps team to clean up legacy code for improved run times to handle larger datasets.
- The main purpose of this refactor is to make the VBA code more efficient with larger datasets.

### OBJECTIVE

- DELIVERABLE 1: Please see attached Refactored VBA Code for the Green-Stocks Analyzer.
- DELIVERABLE 2: Release notes over review of the refactored code compaired to legacy code.

---
---
---

## DELIVERABLE 2

> The following describes issues or faults and the refactored solutions to resolve them or improve functionality.

### Worksheet avalability

#### Issues:
- When user enters a dataset sheet name ("2017", "2018", "2019"...), if the data sheet is not avalable then an error is thrown.

#### Solution 1.0:
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/datasheet_name_avalability.png)

#### Solution Definition 1.0:

1. Count the number of sheets in the workbook in order to aquire sheet count. This is then stored in variable "totalSheets". 
2. If the sheet name exists, then set "worksheetExists" to true and exit for loop. 
3. If sheet name does not exist, then set "worksheeExist" to false.

#### Solution 1.1:

> With the stored bool, the following code was added.
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/datasheet_name_avalability_catch.png)

#### Solution Definition 1.1:

1. If worksheetExist is true, then run the stock-analyzer.
2. if worksheetExist is false, then show message, "Worksheet +variable+ does not exist. Please try again.". Call AllStocksAnalysis() to restart function.


