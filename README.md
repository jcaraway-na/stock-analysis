# Green-Stock Analysis

### BACKGROUND

> Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

The analysis below will provide insight into the following:
- DevOps team to clean up legacy code for improved run times to handle larger datasets.
- The main purpose of this refactor is to make the VBA code more efficient with larger datasets.

### OBJECTIVE

- DELIVERABLE 1: Please see attached Refactored VBA Code for the Green-Stocks Analyzer.
- DELIVERABLE 2: Release notes for review of the refactored code. Will include summaries for areas of improved code.

---
---
---

## DELIVERABLE 2

> The following describes issues or faults and the refactored solutions to resolve them or improve functionality.

### Worksheet User Experience Error Handling

#### Issues 1.0:

1. Input box accepts varchar. should only accept year integers.
2. If user clicks cancel on input box, an error is thrown
3. When user enters a dataset sheet name ("2017", "2018", "2019"...), if the data sheet is not available then an error is thrown.


#### Solution 1.0:
![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/ux_error_handling_res/ux_code_errorhandler1%20(2).png)

#### Solution Definition 1.0:

> The above tries to handle User Experiance issues. 
> 1. Set inputbox to only accept integers": applied a simple object variable validation to only accept integers ONLY.
> 2. If user clicks cancel": added an If input == False Then Exit sub on "cancel" click.
> 3. When user enters dataset sheet name, if sheet is not available": built "sheet detector" function. If a user enters a data sheet name that is not available, 
> then user is prompted, "..."Worksheet " + yearValue + " does not exist. Please try again"...". The Sub routine is then restarted.

---

### Conditional Formatting: Static Range vs. Dynamic Ranging

#### Issues 2.0:

1. Conditional formatting range is hard coded in code behind. This could lead to future issues on usability if a ticker is added to the dataset.

> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/dynamic_indexing_res/original_hardcode_ranging.png)
> Above image shows original code and describes the potintial issue behind code.

#### Solution 2.0:
![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/dynamic_indexing_res/refactored_dynamic_ranging.png)

#### Solution Definition 2.0:

> The above tries to handle User Experiance issues. 
> 1. Set inputbox to only accept integers": applied a simple object variable validation to only accept integers ONLY.
> 2. If user clicks cancel": added an If input == False Then Exit sub on "cancel" click.
> 3. When user enters dataset sheet name, if sheet is not available": built "sheet detector" function. If a user enters a data sheet name that is not available, 
> then user is prompted, "..."Worksheet " + yearValue + " does not exist. Please try again"...". The Sub routine is then restarted.
