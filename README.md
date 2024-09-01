# ExcelQueryCLI

ExcelQueryCLI is a command line tool that allows you to update Excel files using yaml querying.

## You can use ExcelQueryCLI to:

- Update rows in Excel files based on the filter query
- Delete rows in Excel files based on the filter query
- Automate repetitive tasks in Excel files
- Process multiple Excel files in bulk
- Use a simple and intuitive query language to define your operations
- Backup your files before making any changes
- Use parallel processing for faster execution

## Installation

- Download and install .NET 8 SDK from [here](https://dotnet.microsoft.com/download/dotnet/8.0)
- Download the latest release from GitHub and extract the zip file

Portable version can be used in any OS.

Win 64 version is for Windows 64 bit OS and it is a bundled exe file

## Warning

This app still in early development and may contain bugs. Please use it with caution.

Syntax for querying might change with future updates

## Disclaimer

This project uses the EPPlus library for Excel file handling.

EPPlus is licensed under the GNU Library General Public License (LGPL) and is free to use for non-commercial purposes.

For commercial purposes, you need to purchase a license from the [EPPlus website](https://epplussoftware.com/).

## Supported Excel File Formats
It will work on any Excel file that is supported by EPPlus library

## Supported Query File Formats

- YAML
- JSON
- XML

## Query

CLI tool uses YAML, XML, JSON files for querying. The YAML, XML, JSON file should contain the following structure

See more examples in the [examples](Examples) folder

Check [class models](ExcelQueryCLI/Models) in project to have a better understanding of the query structure

You must use "root" element in XML file
### Update Query

At least one update query must be provided for update operation

Complex query with multiple filters

```yaml
source: # You can specify multiple files and directories
  - 'ExcelFile.xlsx'
  - 'ExcelFile2.xlsx'
  - 'Folder\ExcelFiles'
backup: true # Backup files before updating
sheets:
    - name: 'Employees Table' # Name of the sheet
      header_row: '1' # Row number of the header
      start_row: '2' # Row number of the first data row
    - name: 'Salary Table'
    - name: 'Address Table'
query:
  - update: # With single filter
      - column: 'Fullname' # Column name to update
        operator: 'APPEND' # Operator to use SET, ADD, SUBTRACT, MULTIPLY, DIVIDE etc.
        value: 'John Doe' # Value to use for update
      - column: 'Phone' # Column name to update
        operator: 'Replace' # Operator to use SET, ADD, SUBTRACT, MULTIPLY, DIVIDE etc.
        value: '555|>|222' # Value to use for update
    filters: # Filters to apply
      - column: 'NAME'
        compare: 'EQUALS'
        values:
          - 'John'
          - 'Mark'
      - column: 'Salary'
        compare: 'BETWEEN'
        values:
          - '1000-2000'
  - update: # With multiple filters
      column: 'Address' # Column name to update
      operator: 'SET' # Operator to use SET, ADD, SUBTRACT, MULTIPLY, DIVIDE etc.
      value: 'Turkey' # Value to use for update
    filter_merge: 'AND' # Operator to use for multiple filters, it does not have any effect when there is only one filter
    filters: # Filters to apply
      - column: 'NAME' # Multiple filters can be applied
        compare: 'EQUALS'
        values:
          - 'John' # Value to compare
      - column: 'FULLNAME'
        compare: 'EQUALS'
        values:
          - 'Mark'
  - update: # you can use without filter
      - column: 'Salary'
        operator: 'MULTIPLY'
        value: '1.3'
```

Simple query without any filters

```yaml
source:
  - 'ExcelFile.xlsx'
sheets:
    - name: 'Employees Table'
query:
  - update:
      - column: 'Salary'
        operator: 'MULTIPLY'
        value: '1.3'
```

### Delete Query

At least one filter must be provided for delete operation

Simple query

```yaml
source: # Source file or directory path
  - 'ExcelFile.xlsx'
sheets: # Sheet names to be processed
    - name: 'Employees Table' # Sheet name
query:
  - filters: # Filter queries
      - column: 'Department' # Column name to filter
        operator: 'EQUALS' # Operator to use for filter
        values: # Values to use for filter (always using OR operation for compare since otherwise does not make any sense)
          - 'HR'
```

Simple query with multiple filters

```yaml
source: # Source file or directory path
  - 'ExcelFile.xlsx'
sheets: # Sheet names to be processed
    - name: 'Employees Table' # Sheet name
query:
  filter_merge: 'AND' # Filter merge operator (AND, OR) only valid when multiple filters are used
  filters: # Filter queries
    - column: 'Department' # Column name to filter
      operator: 'EQUALS' # Operator to use for filter
      values: # Values to use for filter (always using OR operation for compare since otherwise does not make any sense)
        - 'HR'
    - column: 'Location'
      operator: 'NOT_EQUALS'
      values:
        - 'Turkey'
```

### Query Structure

#### `source` : Source file or directory path

This section specifies the Excel files or directories you want to process.
You can provide multiple file paths or even directories containing Excel files.

#### `backup` : Backup files before updating

Set this to true to create backup copies of your original files before any updates are made.
This provides a safety net in case you need to revert any changes.

#### `sheets` : Sheet names to be processed

Here you define the specific sheets within your Excel files that you want to work with.

For each sheet, provide:

- `name`: The exact name of the sheet in the Excel file.
- `header_row` (optional): The row number containing the column headers. Defaults to the first row if not provided.
- `start_row` (optional): The row number where your actual data begins. Defaults to the second row if not provided.

#### `query` : Query items

This is the core of your configuration, outlining the filtering and update operations you want to perform.

Each query item consists of:

- `update`: Defines the update action. (_only for update queries_)
    - `column`: The name of the column you want to modify.
    - `operator`: The type of update to perform (e.g., SET, ADD, MULTIPLY).
    - `value`: The value to use in the update operation.
- `filter_merge` (optional): Specifies how multiple filters should be combined (AND or OR). Only valid when multiple
  filters are used.
- `filters` (optional): Specifies the conditions to filter rows before applying the update.
    - Each filter includes:
        - `column`: The column to filter on.
        - `compare`: The comparison operator (e.g., EQUALS, CONTAINS).
        - `values`: A list of values to compare against.

### Compare Operators

Used in comparing values in the filter queries

- `EQUALS` : Equals operator
- `NOT_EQUALS` : Not equals operator
- `GREATER_THAN` : Greater than operator
    - Passed value must be a floating number
- `GREATER_THAN_OR_EQUAL` : Greater than or equal operator
    - Passed value must be a floating number
- `LESS_THAN` : Less than operator
    - Passed value must be a floating number
- `LESS_THAN_OR_EQUAL` : Less than or equal operator
    - Passed value must be a floating number
- `CONTAINS` : Contains operator
- `NOT_CONTAINS` : Not contains operator
- `STARTS_WITH` : Starts with operator
- `ENDS_WITH` : Ends with operator
- `BETWEEN` : Between operator
    - You must provide 2 numbers in a single value field separated by a dash (-)
- `NOT_BETWEEN` : Not between operator
    - You must provide 2 numbers in a single value field separated by a dash (-)
- `IS_NULL_OR_BLANK` : Is null or blank operator
    - You can not give any value when using this operator
- `IS_NOT_NULL_OR_BLANK` : Is not null or blank operator
    - You can not give any value when using this operator

### Update Operators

Used in updating values in the update queries

- `SET` : Set operator
- `MULTIPLY` : Multiply operator
    - Passed value must be floating number
- `DIVIDE` : Divide operator
    - Passed value must be floating number
- `ADD` : Add operator
    - Passed value must be floating number
- `SUBTRACT` : Subtract operator
    - Passed value must be floating number
- `APPEND` : Append operator
    - Passed value can not be empty string
- `PREPEND` : Prepend operator
    - Passed value can not be empty string
- `REPLACE` : Replace operator
    - You must provide 2 values in a single value field separated by "|>|" (without quotes)
        - First value is the value to be replaced
        - Second value is the value to replace with

## Update Function

Update rows in Excel file based on the parameters provided

```bash
ExcelQueryCLI.exe update -q <query-file-path> -p <parallelism>
```

### Parameters

- `-q` or `--query` _(required)_
    - The file or directory path to the yaml query file
- `-p` or `--parallelism`
    - The number of parallel threads to use for processing. Default is 1

### Example

```bash
ExcelQueryCLI.exe update -q "update.yaml" -p 4
```

## Delete Function

Delete rows in Excel file based on the parameters provided

```bash
ExcelQueryCLI.exe update -q <query-file-path> -p <parallelism>
```

### Parameters

- `-q` or `--query` _(required)_
    - The file or directory path to the yaml query file
- `-p` or `--parallelism`
    - The number of parallel threads to use for processing. Default is 1

### Example

```bash
ExcelQueryCLI.exe delete -q "delete.yaml" -p 4
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## Changelog

### v2.3
- Fixed XML dictionary serialization issue
- Refactored sheets model which caused syntax model change

### v2.2
- Fixed an issue where JSON and XML query files were not being read correctly

### v2.1

- Added support for JSON and XML query files

### v2.0

- Dumped the OpenXML SDK and switched to EPPlus for better file handling
- Reworked the query language to YAML instead of command parameters
- Added support for multiple filter queries
- Added support for multiple update queries
- Added support for `AND` and `OR` operators in filter queries
- Added support for parallel processing
- Added support for setting header row number
- Added support for setting start row number
- Added 2 new compare operators `IS_NULL_OR_BLANK` and `IS_NOT_NULL_OR_BLANK`
- Better separation of methods
- Improved error handling
- Improved logging
- Improved data type validation and conversion with YAML deserialization
- Improved data validation
- Removed first row update parameter
- Updated syntax to support multiple column updates in query
- Implemented delete functionality
- Added tests project
- Implemented backup feature
- Refactored REPLACE, BETWEEN and NOT_BETWEEN operators

### v1.4

- Added support for directory path
- Added support for multiple `-f` parameters

### v1.3

- Added possibility to update without filter query

### v1.2

- Syntax change
- Added support for multiple column names in single filter query
- Added support for multiple values in single set query

### v1.1

- Removed delete functionality for now
- Bug fixes
- Regex dynamic enum generation

### v1.0

- Initial release