using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Update;

namespace ExcelQueryCLI.Tests;

public sealed class ExcelUpdateQueryTests
{
  [Test]
  public void TestExcelQueryParseXMLText_Complex_Valid() {
    const string json = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <source>ExcelFile2.xlsx</source>
                          <source>Folder\ExcelFiles</source>
                          <backup>true</backup>
                          <sheets name="Employees Table" header_row="1" start_row="2"/>
                          <sheets name="Salary Table"/>
                          <sheets name="Address Table"/>
                          <query>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query>
                          <query>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query>
                          <query>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query>
                        </root>
                        """;
    Assert.DoesNotThrow(() => {
      var excelQuery = ExcelUpdateQuery.ParseXmlText(json);
      Assert.That(excelQuery.Source, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Sheets, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Query, Has.Length.EqualTo(3));
    });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseJSONText_Complex_Valid() {
    const string json = """
                        {
                          "source": [
                            "ExcelFile.xlsx",
                            "ExcelFile2.xlsx",
                            "Folder\\ExcelFiles"
                          ],
                          "backup": true,
                          "sheets": [
                            {
                              "name": "Employees Table",
                              "header_row": "1",
                              "start_row": "2"
                            },
                            {
                              "name": "Salary Table"
                            },
                            {
                              "name": "Address Table"
                            }
                          ],
                          "query": [
                            {
                              "update": [
                                {
                                  "column": "Fullname",
                                  "operator": "APPEND",
                                  "value": "John Doe"
                                }
                              ],
                              "filters": [
                                {
                                  "column": "NAME",
                                  "compare": "EQUALS",
                                  "values": [
                                    "John",
                                    "Mark"
                                  ]
                                }
                              ]
                            },
                            {
                              "update": [
                                {
                                  "column": "Address",
                                  "operator": "SET",
                                  "value": "Turkey"
                                }
                              ],
                              "filter_merge": "AND",
                              "filters": [
                                {
                                  "column": "NAME",
                                  "compare": "EQUALS",
                                  "values": [
                                    "John"
                                  ]
                                },
                                {
                                  "column": "FULLNAME",
                                  "compare": "EQUALS",
                                  "values": [
                                    "Mark"
                                  ]
                                }
                              ]
                            },
                            {
                              "update": [
                                {
                                  "column": "Salary",
                                  "operator": "MULTIPLY",
                                  "value": "1.3"
                                }
                              ]
                            }
                          ]
                        }
                        """;
    Assert.DoesNotThrow(() => {
      var excelQuery = ExcelUpdateQuery.ParseJsonText(json);
      Assert.That(excelQuery.Source, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Sheets, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Query, Has.Length.EqualTo(3));
    });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseYamlText_Complex_Valid() {
    const string yaml = """
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
                          - # With single filter
                            update:
                              -
                                column: 'Fullname' # Column name to update
                                operator: 'APPEND' # Operator to use SET, ADD, SUBTRACT, MULTIPLY, DIVIDE etc.
                                value: 'John Doe' # Value to use for update
                            filters: # Filters to apply
                              -
                                column: 'NAME'
                                compare: 'EQUALS'
                                values: 
                                  - 'John'
                                  - 'Mark'
                          - # With multiple filters
                            update:
                              -
                                column: 'Address' # Column name to update
                                operator: 'SET' # Operator to use SET, ADD, SUBTRACT, MULTIPLY, DIVIDE etc.
                                value: 'Turkey' # Value to use for update
                            filter_merge: 'AND' # Operator to use for multiple filters, it does not have any effect when there is only one filter
                            filters: # Filters to apply
                              - # Multiple filters can be applied
                                column: 'NAME'
                                compare: 'EQUALS'
                                values: 
                                  - 'John' # Value to compare
                              -
                                column: 'FULLNAME'
                                compare: 'EQUALS'
                                values:
                                  - 'Mark'
                          - # you can use without filter
                            update:
                              - 
                                column: 'Salary'
                                operator: 'MULTIPLY'
                                value: '1.3'
                        """;
    Assert.DoesNotThrow(() => {
      var excelQuery = ExcelUpdateQuery.ParseYamlText(yaml);
      Assert.That(excelQuery.Source, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Sheets, Has.Length.EqualTo(3));
      Assert.That(excelQuery.Query, Has.Length.EqualTo(3));
    });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseYamlText_Simple_Valid() {
    const string yaml = """
                        source:
                          - 'ExcelFile.xlsx'
                        sheets:
                          - name: 'Employees Table'
                        query:
                          - update:
                              - column: 'Salary'
                                operator: 'MULTIPLY'
                                value: '1.3'
                        """;
    Assert.DoesNotThrow(() => {
      var excelQuery = ExcelUpdateQuery.ParseYamlText(yaml);
      Assert.That(excelQuery.Source, Has.Length.EqualTo(1));
      Assert.That(excelQuery.Sheets, Has.Length.EqualTo(1));
      Assert.That(excelQuery.Query, Has.Length.EqualTo(1));
      foreach (var VARIABLE in excelQuery.Query) Assert.That(VARIABLE.Filters, Is.Null);
    });
    Assert.Pass();
  }


  [Test]
  public void TestExcelQueryParseSimpleYamlText_Invalid() {
    const string yaml = """
                        source:
                          - 'ExcelFile.xlsx'
                        sheets:
                          - name: 'Employees Table'
                        query:
                          - update:
                              - column: 'Salary'
                                operator: 'MULTIPLY'
                                value: '1.3'  
                              - column: 'Salary'
                                operator: 'DIVIDE'
                                value: '3'
                        """;
    Assert.Throws<ArgumentException>(() => { _ = ExcelUpdateQuery.ParseYamlText(yaml); }, "Update column names must be unique");
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleYamlText_Invalid2() {
    const string yaml = """
                        source:
                          - 'ExcelFile.xlsx'
                        sheets:
                          - name: 'Employees Table'
                        query:
                          - update:
                            - column: 'Salary'
                              operator: 'MULTIPLY'
                              value: 'qeqew'
                        """;
    Assert.Throws<ArgumentException>(() => {
                                       var a = ExcelUpdateQuery.ParseYamlText(yaml);
                                       Console.WriteLine();
                                     },
                                     "MULTIPLY operator requires a number value");
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleYamlText_Invalid3() {
    const string yaml = """
                        source:
                          - 'ExcelFile.xlsx'
                        sheets:
                          - name: 'Employees Table'
                        query:
                          - update:
                            - column: 'Salary'
                              operator: 'APPEND'
                              value: ''
                        """;
    Assert.Throws<ArgumentException>(() => { _ = ExcelUpdateQuery.ParseYamlText(yaml); }, "APPEND operator requires a non-empty value");
    Assert.Pass();
  }
}