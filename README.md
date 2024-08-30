﻿# ExcelQueryCLI

ExcelQueryCLI is a command line tool that allows you to query Excel files using simple querying.

## You can use ExcelQueryCLI to:

- Update rows in excel file by filtering columns

## Installation

Download the latest release from github and extract the zip file

## Warning
This app still in early development and may contain bugs. Please use it with caution.

Syntax for querying might change with future updates

## Query

Simple query language contains 3 important values :

- Column Name
- Operator
- Value

```bash
"('<column-name>') <operator> ('<value>')"
"('<column-name>' OR '<column-name>') <operator> ('<value>' OR '<value>')"
```

The column name and value should be enclosed in single quotes and parenthesis like in the example above~~~~

### Query Types

- Filter Query : Used to filter rows in excel file
- Set Query : Used to update rows in excel file

### Filter Query Operators

- `EQUALS` : Equals operator
- `NOT_EQUALS` : Not equals operator
- `GREATER_THAN` : Greater than operator
- `GREATER_THAN_OR_EQUAL` : Greater than or equal operator
- `LESS_THAN` : Less than operator
- `LESS_THAN_OR_EQUAL` : Less than or equal operator
- `CONTAINS` : Contains operator
- `NOT_CONTAINS` : Not contains operator
- `STARTS_WITH` : Starts with operator
- `ENDS_WITH` : Ends with operator
- `BETWEEN` : Between operator
- `NOT_BETWEEN` : Not between operator

BETWEEN and NOT_BETWEEN operators require the values to be a list of values separated by (<>)
 
### Set Query Operators

- `SET` : Set operator
- `MULTIPLY` : Multiply operator
- `DIVIDE` : Divide operator
- `ADD` : Add operator
- `SUBTRACT` : Subtract operator
- `APPEND` : Append operator
- `PREPEND` : Prepend operator
- `REPLACE` : Replace operator

MULTIPLY, DIVIDE, ADD and SUBTRACT operators require the value of column and set value to be a floating number
## Update

Update rows in excel file based on the parameters provided

```bash
ExcelQueryCLI.exe update -f <file> -s <sheet> --filter-query <filter-query> --set-query <set-query> --only-first <only-first>
```

### Parameters

- `-f` or `--file` _(required)_
  - The file or directory path to the excel file
  - If a directory path is provided, all excel files in the directory will be processed
  - If a file path is provided, only that file will be processed
  - You can pass multiple file paths and directory paths by using -f multiple times
- `-s` or `--sheet` _(required)_
  - The name of the sheet in the excel file
- `--filter-query` 
  - The filter query to filter the rows to be updated
  - If no filter query is provided, all rows in the sheet will be updated
  - You can pass multiple filter query parameters to filter rows by OR operation
- `--set-query` _(required)_
  - The set query to update the rows
  - You can pass multiple set query parameters to update the rows
- `--only-first`
  - If set, only the first row that matches the filter query will be updated
- `--header-row-index`
  - The row number of the header row in the sheet. Default is 1

### Example

**Simple query**
```bash
ExcelQueryCLI.exe update -f "sample.xlsx" -s "Sheet1" --filter-query "('Name') EQUALS ('John Doe')" --set-query "('Age') SET ('30')"
```

**Complex query**
```bash
ExcelQueryCLI.exe update -f "sample.xlsx" -f "sample2.xlsx"  -f "D:\\SampleDirectory" -s "Sheet1" --filter-query "('Name' OR 'Surname' OR 'Fullname') NOT_EQUALS ('John' OR 'Mark' OR 'Justin')" --set-query "('Age') SET ('30')" --set-query "('UserPermission') SET ('3')"
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details