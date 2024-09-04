# ExcelQueryCLI

ExcelQueryCLI is a command line tool that allows you to update Excel files using XML, JSON, YAML querying.

You create a query file that defines the operations you want to perform on your Excel files, such as filters.

The tool reads the query file and applies the operations to the Excel files, updating the rows based on the filter conditions.

You can use this tool to automate repetitive tasks in Excel files, such as updating rows based on certain conditions or deleting rows that meet specific criteria.

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

Breaking changes could be introduced in pre-release.

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

### Query Rules

#### Handled Gracefully
- Do not pass duplicated source paths (Handled gracefully)
- Do not pass duplicated sheet names in root scope (Handled gracefully)
- Do not pass duplicated sheet names in query scope (Handled gracefully)
- When using values_def_key the key must not contain any spaces (Handled gracefully)
- All column names is trimmed before usage

#### Throws Error
- Do not pass duplicated column names in update scope 
- Do not pass duplicated column names in filter scope 
- Do pass either global sheet name or sheet name in every query scope 
- Do not pass values in filter scope for IS_NULL_OR_BLANK and IS_NOT_NULL_OR_BLANK operators 
- Do pass values separated by '-' in filter scope for BETWEEN and NOT_BETWEEN operators 
- Do pass 2 values separated by '|>|' in update scope for REPLACE operator 
- Do not pass empty string in value field for APPEND and PREPEND operators 
- You must provide multiple filters when using filter_merge 
- You can pass filter_merge when there is not multiple filters passed 


### Update Query

At least one update query must be provided for update operation

Complex query

```xml
<?xml version="1.0" encoding="utf-8" ?>
<root>
  <source>ExcelFile.xlsx</source>
  <source>ExcelFile2.xlsx</source>
  <source>Folder\ExcelFiles</source>
  <backup>true</backup>
  <sheets name="Employees Table" header_row="1" start_row="2"/>
  <values_def key="NAMES">
    <values>John</values>
    <values>Mark</values>
    <values>Justin</values>
  </values_def>
  <query>
    <sheets name="Address Table"/>
    <update column="Fullname" operator="APPEND" value="John Doe"/>
    <update column="Department" operator="SET" value="HR"/>
    <filters column="NAME" compare="EQUALS">
      <values_def_key>NAMES</values_def_key>
    </filters>
  </query>
  <query>
    <sheets name="Salary Table"/>
    <update column="Address" operator="SET" value="Turkey"/>
    <filter_merge>AND</filter_merge>
    <filters column="NAME" compare="EQUALS">
      <values_def_key>NAMES</values_def_key>
    </filters>
    <filters column="FULLNAME" compare="EQUALS">
      <values_def_key>NAMES</values_def_key>
      <values>Ella</values>
      <values>Lawrance</values>
    </filters>
  </query>
  <query>
    <update column="Salary" operator="MULTIPLY" value="1.3"/>
  </query>
</root>
```

Simple query

```xml
<?xml version="1.0" encoding="utf-8" ?>
<root>
  <source>ExcelFile.xlsx</source>
  <sheets name="Employees Table" header_row="1" start_row="2"/>
  <query>
    <update column="Fullname" operator="APPEND" value="John Doe"/>
    <update column="Department" operator="SET" value="HR"/>
  </query>
</root>
```

### Delete Query

At least one filter must be provided for delete operation

Complex query

```xml
<?xml version="1.0" encoding="utf-8" ?>
<root>
  <source>ExcelFile.xlsx</source>
  <source>ExcelFile2.xlsx</source>
  <source>Folder\ExcelFiles</source>
  <backup>true</backup>
  <sheets name="Employees Table" header_row="1" start_row="2"/>
  <values_def key="NAMES">
    <values>John</values>
    <values>Mark</values>
    <values>Justin</values>
  </values_def>
  <query>
    <sheets name="Address Table"/>
    <filters column="NAME" compare="EQUALS">
      <values_def_key>NAMES</values_def_key>
    </filters>
  </query>
  <query>
    <sheets name="Salary Table"/>
    <filter_merge>AND</filter_merge>
    <filters column="NAME" compare="EQUALS">
      <values_def_key>NAMES</values_def_key>
      <values>Ella</values>
    </filters>
    <filters column="FULLNAME" compare="EQUALS">
      <values>Mark</values>
    </filters>
  </query>
</root>
```

Simple query

```xml
<?xml version="1.0" encoding="utf-8" ?>
<root>
  <source>ExcelFile.xlsx</source>
  <sheets name="Employees Table" header_row="1" start_row="2"/>
  <query>
    <filters column="NAME" compare="EQUALS">
      <values>John</values>
    </filters>
  </query>
</root>
```

### Query Structure

#### `source` : Source file or directory path

This section specifies the Excel files or directories you want to process.

You can provide multiple file paths or even directories containing Excel files.

When passing directories it will process all Excel files in the directory. If duplicate files are passed it will be handled gracefully.

#### `backup` : Backup files before updating

Set this to true to create backup copies of your original files before any updates are made.

This provides a safety net in case you need to revert any changes.

It creates a backup folder in the same directory as the app and saves the backup files there with a timestamp.

#### `sheets` : Sheet names to be processed

Here you define the specific sheets within your Excel files that you want to work with.

For each sheet, provide:

- `name`: The exact name of the sheet in the Excel file.
- `header_row` (optional): The row number containing the column headers. Defaults to the first row if not provided.
- `start_row` (optional): The row number where your actual data begins. Defaults to the second row if not provided.

Header row can not be greater than start row

#### `values_def` : Define values to be reused in query filters

Here you can define reusable values to be used in query filters.

Define it like this in root:

```xml
<values_def key="my_custom_key">
  <values>45</values>
  <values>346</values>
  <values>634</values>
</values_def>
```

Then you can use it in your query like this:

```xml
<filters column="ID" compare="EQUALS">
  <values_def_key>HR_IDS</values_def_key>
  <values_def_key>DEV_IDS</values_def_key>
</filters>
```

**Rules:**

- Keys must be unique
- Key name should not contain any spaces or special characters
- You can use multiple definition keys in a single filter query which will concatenate the values

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
        - `values_def_key`: A key to reference the values defined in the `values_def` section.

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
    - Passed value can not be empty string
- `NOT_CONTAINS` : Not contains operator
    - Passed value can not be empty string
- `STARTS_WITH` : Starts with operator
    - Passed value can not be empty string
- `ENDS_WITH` : Ends with operator
    - Passed value can not be empty string
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

## Usage

```bash
Usage: ExcelQueryCLI [command] [query-file-path] [options]
```

```bash
Commands:
  update    Update rows in Excel file
  delete    Delete rows in Excel file
```

```bash
Options:
  -l, --log-level <LogEventLevel>    Log level (Default: Information) (Allowed values: Verbose, Debug, Information, Warning, Error, Fatal)
  -c, --commercial                   Use commercial license
  -p, --parallel-threads <Byte>      Number of parallel threads (Default: 1)
  -h, --help                         Show help message
```

## Update Function

Update rows in Excel file based on the parameters provided

```bash
ExcelQueryCLI.exe update <query-file-path> -p <parallelism> -l <log-level> -c <commercial>
```

Example usage

```bash
ExcelQueryCLI.exe update "update.xml" -p 4 -l Debug -c true
```

## Delete Function

Delete rows in Excel file based on the parameters provided

```bash
ExcelQueryCLI.exe delete -q <query-file-path> -p <parallelism> -l <log-level> -c <commercial>
```

Example usage

```bash
ExcelQueryCLI.exe delete -q "delete.xml" -p 4 -l Debug -c true
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details

## Changelog

### v2.5

- Added possibility to pass sheets path inside query element. This will be concatenated with the root sheet paths
- Added option to set log level via '-l' or '--log-level' parameter
- Added option to set commercial usage via '-c' or '--commercial' parameter (if enabled you need to use EPPlus license
  file)
- Query file no longer needs parameter name you can pass it as argument after 'delete' or 'update' command (this is a
  breaking change)
- Fixed an issue where duplicated source files can be passed when passing directories
- Update cell function now checks if old value is same as new value before updating resulting in correct update count
- Added support for defining and reusing values in query files (values_def element)

### v2.4

- Fixed an issue where JSON property name was not working correctly
- Fixed an issue where XML parsing was not working correctly due to wrong attribute usage
- Fixed an issue where query validation were not working for JSON and XML files
- Added more indepth tests

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