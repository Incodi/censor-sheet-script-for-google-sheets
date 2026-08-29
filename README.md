# censor-sheet-script-for-google-sheets
App script to censor a Sheet in the active spreadsheet in a Google Sheets. 

This is good for small data sets, but does not scale well as it duplicates the target sheet to make a censored one.

# How this works

It checks rows in column E of the target sheet if it has any words that are in column A in the censored list. If it does have an offending word, it turns the entire cell into "[Censored]"
