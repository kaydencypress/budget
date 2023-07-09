# budget

Generates an Excel spreadsheet summarizing monthly expenses by category vs budget. Assigns category information based on the source data where possible, and also supports custom user-defined mappings by transaction name.

Before running this script, export credit card and/or bank account transactions in CSV format and save to the /import directory without renaming the files. Chase credit card and Wells Fargo bank accounts are supported; other financial institutions have not been tested.

usage: budget_report.py [-h] [-b] [-c]

options:
  -h, --help        show this help message and exit
  -b, --budget      review and modify budget
  -c, --categorize  categorize unmapped transactions

Example output (censored):
![example_output](https://github.com/kaydencypress/budget/assets/127451126/fe959ad3-7847-4f30-acd1-1df55ee917a4)
