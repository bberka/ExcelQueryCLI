
# ExcelQueryCLI
ExcelQueryCLI is a command line tool that allows you to query Excel files using simple querying.

## You can use ExcelQueryCLI to:
- Update rows in excel file 
- Delete rows in excel file 

## Installation
Download the latest release from github and extract the zip file

## Query
Simple query language contains 3 important values :
- Column Name
- Operator
- Value

```bash
"'<column-name>' <operator> '<value>'"
```

**Note:** The column name and value should be enclosed in single quotes

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
- `IN` : In operator 
- `NOT_IN` : Not in operator
- `BETWEEN` : Between operator
- `NOT_BETWEEN` : Not between operator

**Note:** IN, NOT_IN, BETWEEN and NOT_BETWEEN operators require the values to be a list of values separated by (|) pipe character


### Set Query Operators
- `SET` : Set operator
- `MULTIPLY` : Multiply operator
- `DIVIDE` : Divide operator
- `ADD` : Add operator
- `SUBTRACT` : Subtract operator
- `APPEND` : Append operator
- `PREPEND` : Prepend operator
- `REPLACE` : Replace operator

**Note:** MULTIPLY, DIVIDE, ADD and SUBTRACT operators require the value to be a floating number


## Delete 
Delete rows in excel file based on the parameters provided
```bash
ExcelQueryCLI.exe delete -f <file> -s <sheet> --filter-query <filter-query> --only-first <only-first>
```
### Parameters
- `-f` or `--file` : The path to the excel file
- `-s` or `--sheet` : The name of the sheet in the excel file
- `--filter-query` : The filter query to filter the rows to be deleted
- `--only-first` : If set, only the first row that matches the filter query will be deleted

### Example
```bash
ExcelQueryCLI.exe delete -f "sample.xlsx" -s "Sheet1" --filter-query "'Name' EQUALS 'John Doe'"
```

## Update
Update rows in excel file based on the parameters provided
```bash
ExcelQueryCLI.exe update -f <file> -s <sheet> --filter-query <filter-query> --set-query <set-query> --only-first <only-first>
```
### Parameters
- `-f` or `--file` : The path to the excel file
- `-s` or `--sheet` : The name of the sheet in the excel file
- `--filter-query` : The filter query to filter the rows to be updated
- `--set-query` : The set query to update the rows
- `--only-first` : If set, only the first row that matches the filter query will be updated

### Example
```bash
ExcelQueryCLI.exe update -f "sample.xlsx" -s "Sheet1" --filter-query "'Name' EQUALS 'John Doe'" --set-query "'Age' SET '30'"
```
```bash
ExcelQueryCLI.exe update -f "sample.xlsx" -s "Sheet1" --filter-query "'Name' IN 'John|Mark|Justin'" --set-query "'Age' MULTIPLY '2'"
```

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details