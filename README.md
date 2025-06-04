
### 4june25_1206m
- `auto_id_generate_regenerate_on_any_change.js` ✅

    - **Description**: 
        - Auto-updates IDs in Column A for sheet name ending with word `-Expenses` after any change   (edit, row insert, row delete, paste). Includes an `Expense Tools` menu for manual ID regeneration wheenver you want.

    - **Features**:
      - Generates unique IDs in Column A based on non-empty Column B entries.
      - ID format: `[XX]-[number]` (e.g., `AD-1` for `Admin-Expenses` sheet).
      - Updates IDs dynamically on any spreadsheet change via `onChange` trigger.
      - Supports manual ID regeneration through the "Expense Tools" menu.
      - Syncs IDs with non-empty rows in Column B.

    - **Setup**:
      1. Open the script editor in Google Sheets (`Extensions > Apps Script`).
      2. Paste the script content.
      3. Add a trigger:
         - Click the clock icon (Triggers).
         - Select `+ Add Trigger`.
         - Choose Choose which function to run - `onChange`
         - Choose which deployment should run - `Head`
         - Select event source - `From spreadsheet`
         - Select event type - `On change`
         - Click on `Save` to save trigger.
      4. Save the script and refresh the sheet to access the "Expense Tools" menu.

    - **Notes**:
      - Assumes the first row is a header; processes data from row 2.
      - Works only on sheets ending with "-Expenses".
      - Shows an alert if manual regeneration is attempted on a non-compatible sheet.

---

- `auto_id_generate_regenerate_on_edit_and_add_row.js` ⭕
    - Auto-generates unique IDs in Column A based on entries in Column B for sheets ending with `-Expenses`.
    - IDs use the format [XX]-[number], where XX is the sheet prefix (e.g., AD-1).
    - Adds a custom `Expense Tools` menu on top in G-sheet to manually regenerate all IDs.
    - Efficiently updates all IDs at once and ensures IDs stay in sync with non-empty rows in Column B.
    - ID will be generated and modified automatically in below case
        - editing column B ( i.e ususally a Expense date) - very normal case 😆
        - adding a row in between, then you add date in that row so it will be auto generated and changed also for adjacent row id's.
        - `NOT WORKS IF any row or rows get deleted in Bulk, Adjacent Ids will not be updated 😑`

---

- `auto_generate_id_on_edit.js` - ⁉️ have some issues 
    - Auto-generates unique IDs in Column A when a value is entered in Column B, for sheet name ending with `-Expenses`.
    - ID format: [XX]-[number], where XX is the sheet prefix (e.g., AD-1).
    - IDs update dynamically and clear if Column B is empty.

### 15may25_846am
- `pre_post_hotfix_release_email_merged.js`
    - Adds 6 different buttons for drafting and sending emails for `pre`, `post` and `hotfix` release communication over email.
    - Subject will have timestamp
    - subject will support emojis
    - Confuguration Options
        - users who can see custom menu button.
        - subject for each email.
        - pretable content for each email.
        - posttable content is same now for all email, but can be configured.
        - whether to send visible columns only or hidden columns as well.
        
### 2may25_925pm
- `wrapText_in_edited_cell.js`
    - Automatically enables "Wrap Text" formatting for any cell that is edited and not left empty, ensuring all content is visible.
    - On opening the sheet, sets the active cell to 10 rows below the last filled cell in column C (if any), for quick navigation.
    - Checks for duplicate Bug IDs in column A (when a cell in column A is edited):
        - If a duplicate is found, shows an alert, highlights the cell, and marks it as "Duplicate".
    - Designed for sheets where column A tracks unique Bug IDs and column C is actively updated.
### 1may25_925pm
- `sendTableEmail_final_1may25_738pm.js`
    - This script adds a custom menu to a Google Spreadsheet that allows authorized users to send an email containing a formatted HTML table of data from the active sheet. 
    - The email includes options for CC and BCC, and it can be configured to send only visible columns or all columns from the sheet.




# ***********************************************************
# Data Summary Using Query

### Department wise Expense
##### Here `2024 - ReBorn` is other sheet name,where the data is
##### F - Department Column
##### D - Amount Column
```
=QUERY('2024 - ReBorn'!A2:F, "SELECT F, SUM(D) WHERE F IS NOT NULL GROUP BY F ORDER BY SUM(D) DESC LABEL F 'Department', SUM(D) 'Value'", 0)
```
#### Results
| Department           | Value   |   |   |   |
|----------------------|---------|---|---|---|
| Administration       | 1095332 |   |   |   |
| Branding & Marketing | 919514  |   |   |   |
| HR                   | 545940  |   |   |   |

#

### Month wise Expense (Asc order, NOT ACTUAL MONTH ORDER)
##### Here `2024 - ReBorn` is other sheet name,where the data is
##### A - Date Column
##### D - Amount Column
##### F - Department Column
```
=QUERY(ARRAYFORMULA({TEXT('2024 - ReBorn'!A3:A, "mmmm"), '2024 - ReBorn'!D3:D}), "SELECT Col1, SUM(Col2) WHERE Col1 IS NOT NULL GROUP BY Col1 LABEL Col1 'Month', SUM(Col2) 'Total Expenses'", 0)
```
#### Results
| Month     | Total Expenses |
|-----------|----------------|
| April     |          35362 |
| August    |         161527 |
| December  |         345489 |
| February  |      105553.12 |
| January   |      131197.33 |
| July      |         249395 |
| June      |         418768 |
| March     |       92020.06 |
| May       |      240757.38 |
| November  |       717985.8 |
| October   |         243102 |
| September |       95691.14 |

#

### Expense Total by Month Number (Number starting from 0 for january and 11 for December)
##### Here `2024` is other sheet name, where the data is
##### A - Date Column
##### D - Amount Column
```
=QUERY('2024'!A2:E, "SELECT MONTH(A), SUM(D) WHERE A IS NOT NULL GROUP BY MONTH(A) ORDER BY MONTH(A) LABEL MONTH(A) 'Month', SUM(D) 'Total Expense'", 0)
```
#### Results
| Month | Total Expense |
|-------|---------------|
| 0     | 133697        |
| 1     | 105553        |
| 2     | 92020         |
| 3     | 35362         |
| 4     | 240757        |
| 5     | 418768        |
| 6     | 249395        |
| 7     | 161527        |
| 8     | 95691         |
| 9     | 243102        |
| 10    | 717986        |
| 11    | 345489        |

#

### Expense Total by Month Number (Number starting from 1 for january and 12 for December)
##### Here `2024` is other sheet name, where the data is
##### A - Date Column
##### D - Amount Column
```
=QUERY('2024'!A2:D, "SELECT MONTH(A) + 1, SUM(D) WHERE A IS NOT NULL GROUP BY MONTH(A) + 1 LABEL MONTH(A) + 1 'Month', SUM(D) 'Total Expenses'", 0)
```
#### Results
| Month | Total Expense |
|-------|---------------|
| 1     | 133697        |
| 1     | 105553        |
| 2     | 92020         |
| 3     | 35362         |
| 4     | 240757        |
| 5     | 418768        |
| 6     | 249395        |
| 7     | 161527        |
| 8     | 95691         |
| 9     | 243102        |
| 10    | 717986        |
| 12    | 345489        |

# *************************************************************************
# Array Formulas
###### Ctrl+Shift+Enter (or Cmd+Shift+Enter on Mac) while editing the formula in the formula bar; this will automatically add the "ARRAYFORMULA" function to the beginning of your formula, making it an array formula.

### Prevents the #VALUE! error by ensuring that the SPLIT function is only called when there is a non-empty value in column E.
#### Formula 1
###### Simple but it's not handling non-visible characters
```
=ARRAYFORMULA(IF(A3:A="", "", IF(E3:E<>"", SPLIT(E3:E, " → ", 0), "")))
```

#### Formula 2
###### Slightly more complex but More robust against non-visible characters (like spaces) since it checks the actual length of the content.
```
=ARRAYFORMULA(IF(A3:A="", "", IF(LEN(E3:E), SPLIT(E3:E, " → ", 0), "")))
```

### Pressing Ctrl+Shift+Enter while editing a formula will automatically add ARRAYFORMULA( to the beginning of the formula. 
```
=ArrayFormula(IF(A3:A="",,SPLIT(E3:E," → ",0)))
```

# **************************************************************************
# INDEX
##### The INDEX function retrieves the value from a specific position within a range or array. Its syntax is:
```
INDEX(array, row_num, [column_num])
```
array: The range or array from which to retrieve data.
row_num: The row number in the array.
column_num (optional): The column number in the array.

Example :
=INDEX(A1:B5, 2, 1)
Above formula returns the value in the second row and first column of the range A1:B5.

### Match with Regex + if no match Found then shows Custom Error Message with value you are Trying to Match due to `IFERROR` function
###### If A12 contains "HR", the message will read "NO Categories found for the department HR".
```
=IFERROR(INDEX(B2:E10, 0, MATCH(TRUE, REGEXMATCH(B1:E1, "^" & A12 & "(?:\s|\n)?Categories$"), 0)), "NO Categories found for the department " & A12)
```

#

### Match with Regex (it will match "HR Categories" OR "HR\nCategories" OR "HRCategories" and 'Categories' word at last)
```
=IFERROR(INDEX(B2:E10, 0, MATCH(TRUE, REGEXMATCH(B1:E1, "^" & A12 & "(?:\s|\n)?Categories$"), 0)), "No Match Found")
```

#

### Match with Regex (it will match "HR Categories" , space and 'Categories' word at last)
```
=INDEX(B2:E10, 0,MATCH(TRUE, REGEXMATCH(B1:E1, "^" & A12 & " Categories$"), 0))
```

#

### Formula to Partially Match the Value of A12 from range B1:E1, As 1 is last arg of match function
```
=INDEX(B2:E10, 0,MATCH(A12,B1:E1,1))
```

#

### Formula to Exactly match the Value of A12 from range B1:E1, As 0 is last arg of match function
```
=INDEX(B2:E10, 0,MATCH(A12,B1:E1,0))
```

# **************************************************************************
# Utilities
### Split the value of E3 with the Delimeter (2nd argument) and expand results in adjacent columns
##### E - Combination Column
##### F - Department Column
##### G - Category Column
##### H - Sub-Category Column
```
SPLIT(E3," → ",0)
```
#### Results
|              Combination             | Department |        Category       | Sub-Category |
|:------------------------------------:|:----------:|:---------------------:|:------------:|
| HR → Rewards & Recognition → Rewards | HR         | Rewards & Recognition | Rewards      |

This can be used with `ArrayFormula` for best usage to give below type of Data
```
=ARRAYFORMULA(IF(A2:A="", "", IF(E2:E<>"", SPLIT(E2:E, " → ", 0), "")))
```
#### Results
| Combination                                                     | Department     | Category                | Sub-Category         |
|-----------------------------------------------------------------|----------------|-------------------------|----------------------|
| HR → Rewards & Recognition → Rewards                            | HR             | Rewards & Recognition   | Rewards              |
| Administration → Office Supplies → Grocery                      | Administration | Office Supplies         | Grocery              |
| Administration → Operational Maintenance → Building Maintenance | Administration | Operational Maintenance | Building Maintenance |
| Administration → Operational Maintenance → Electricity Bill     | Administration | Operational Maintenance | Electricity Bill     |
| IS → Hardware → Cover / Protector                               | IS             | Hardware                | Cover / Protector    |

# ***********************************************************
# Conditional Formatting

### Checks if the value in cell B2 is a number and if the value in cell K2 is blank, 
###### Used to set K2 cell bg to red if the payment date K2 is empty and Trascation date B2 is there
```
=AND(ISNUMBER(B2), ISBLANK(K2))
```

### This formula checks if the day of the week for each date in column A is either Sunday (1) or Saturday (7), Used for conditional Formatting
```
=OR(WEEKDAY(A1)=1, WEEKDAY(A1)=7)
```

## IF

### This formula checks if the day of the week for each date in column A is either Sunday (1) or Saturday (7), Used for conditional Formatting
#### Syntax
```
IF(condition to evaluate, value_if_true, value_if_false)
```
#### Example
`=IF(A1 > 10, "Over 10", "10 or less")`

## IFERROR

### This formula checks if the day of the week for each date in column A is either Sunday (1) or Saturday (7), Used for conditional Formatting
#### Syntax
```
IFERROR(value, value_if_error)
```
#### Example
`=IFERROR(INDEX(B2:B10, MATCH(D1, A2:A10, 0)), "Not Found")`
###### Above formula searches for the value in D1 within the range A2:A10. If found, it returns the corresponding value from B2:B10. If an error occurs (e.g., the value isn't found), it returns "Not Found".

# ***********************************************************
# References for making this type of Readme.md file Online
1. https://markdownlivepreview.com/ - Preview md file online

2. https://www.tablesgenerator.com/markdown_tables - Copy Paste Table / CSV data online and convert to md file data 

### Others
3. https://10015.io/tools/code-to-image-converter - Code to Image Converter

