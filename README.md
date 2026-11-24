# Beginner's Guide to Essential Excel Formulas

## Golden Rules
* **Always start with `=`**: Every formula must begin with an equals sign (`=`), otherwise Excel treats it as plain text.
* **Comma (`,`) vs. Semicolon (`;`)**: Depending on your Excel version or system region settings (such as European locales or some LibreOffice/ONLYOFFICE setups), you may need to use `;` instead of `,` between arguments (e.g., `=SUM(A1; A5)`).

## Basic Math & Stats (A2:A4 means- from A4 block to A4 block)
* **`SUM`**: Adds all numbers in a range (`=SUM(A2:A10)`).
* **`AVERAGE` (AVG)**: Finds the average of numbers in a range (`=AVERAGE(A2:A10)`).
* **`MAX`**: Finds the largest number in a range (`=MAX(A2:A10)`).
* **`MIN`**: Finds the smallest number in a range (`=MIN(A2:A10)`).
* **`COUNT`**: Counts how many cells contain numbers (`=COUNT(A2:A10)`).

## Logic & Conditions
* **`IF`**: Checks a condition and returns one result if true, and another if false (`=IF(B2>50, "Pass", "Fail")`).
* **`COUNTIF`**: Counts cells that match **one** condition (`=COUNTIF(A2:A10, "Tashkent")`).
* **`COUNTIFS`**: Counts rows that match **multiple** conditions (`=COUNTIFS(A2:A10, "Tashkent", B2:B10, ">2")`).
* **`SUMIF`**: Adds numbers that match **one** condition (`=SUMIF(A2:A10, "Tashkent", C2:C10)`).
* **`SUMIFS`**: Adds numbers that match **multiple** conditions — sum range comes first (`=SUMIFS(C2:C10, A2:A10, "Tashkent", B2:B10, ">2")`).

## Lookup
* **`XLOOKUP`**: Searches for a value in one column and returns the matching value from another column (`=XLOOKUP(D2, Products!$A$2:$A$7, Products!$B$2:$B$7)`).
