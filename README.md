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

### Worksheet User Experience Error Handling

#### Issues:

1. Input box accepts varchar. should only accept year integers.
2. If user clicks cancel on input box, an error is thrown
3. When user enters a dataset sheet name ("2017", "2018", "2019"...), if the data sheet is not available then an error is thrown.

#### Solution 1.0:
![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/datasheet_name_avalability1.png)

#### Solution Definition 1.0:

> The above tries to handle User Experiance issues. 
> 1. "set inputbox to only accept integers": applied a simple object variable validation to only accept integers ONLY.
> 2. "if user clicks cancel": applied If input == False Then Exit sub on cancel click.
> 3. "when user enters dataset sheet name, if sheet is not available": built "sheet detector" function. if a user enters a data sheet name that is not available, 
> then user is prompted, "..."Worksheet " + yearValue + " does not exist. Please try again"...". the Sub routine is then restarted.

#### Solution 1.1:

> With the stored bool, the following code was added.
>
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/datasheet_name_avalability_catch.png)

#### Solution Definition 1.1:

1. If worksheetExist is true, then run the stock-analyzer.
2. if worksheetExist is false, then show message, "Worksheet +variable+ does not exist. Please try again.". Call AllStocksAnalysis() to restart function.

#### Refactor Result 1.1:
> Notification window pops-up informing the user that the entered sheet name does not exist. Please try again. The macro is then restarted; bringing the user back
> to the enter year control.
>
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/worksheet_notexist.png)

### Conditional Formatting

#### Issues:
- Conditional formatting macro is set to a constant range. If a user adds a ticker symbole to the ticker array, the conditional formatting macro will ignore the added symbole. The following refactor will address this issue.

> Original Code:
> 
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/conditional_res/original_conditional_rows.png)

#### Solution 2.0:

> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/conditional_res/refactor_conditional_rows.png)

#### Solution Definition 2.1:

1. Dynamically count rows that need conditional formatting.
2. There is no need for constants startRow and endRow variables.

> Once Dynamic range is stored, the stored vaiables are then applied to the conditional formatting For Loop.
>
> ![This is an image](https://github.com/jcaraway-na/stock-analysis/blob/main/resources/refactor_res/conditional_res/conditional_complete_refactor.png)

#### Solution Definition 2.1.1:

1. Dynamically count rows that need conditional formatting.
2. There is no need for constants startRow and endRow variables.
