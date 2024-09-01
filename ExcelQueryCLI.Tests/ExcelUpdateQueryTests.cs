using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Update;

namespace ExcelQueryCLI.Tests;

public sealed class ExcelUpdateQueryTests
{
  [Test]
  public void TestExcelQueryParseYamlText_Complex_Valid() {
    const string yaml = """
                        source: # You can specify multiple files and directories
                          - 'ExcelFile.xlsx'
                          - 'ExcelFile2.xlsx'
                          - 'Folder\ExcelFiles'
                        backup: true # Backup files before updating
                        sheets:
                          employee:
                            name: 'Employees Table' # Name of the sheet
                            header_row: '1' # Row number of the header
                            start_row: '2' # Row number of the first data row
                          salary:
                            name: 'Salary Table'
                          address:
                            name: 'Address Table'
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
      Assert.That(excelQuery.Sheets, Has.Count.EqualTo(3));
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
                          employee:
                            name: 'Employees Table'
                        query:
                          - update:
                              - column: 'Salary'
                                operator: 'MULTIPLY'
                                value: '1.3'
                        """;
    Assert.DoesNotThrow(() => {
      var excelQuery = ExcelUpdateQuery.ParseYamlText(yaml);
      Assert.That(excelQuery.Source, Has.Length.EqualTo(1));
      Assert.That(excelQuery.Sheets, Has.Count.EqualTo(1));
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
                          employee:
                            name: 'Employees Table'
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
                          employee:
                            name: 'Employees Table'
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
                          employee:
                            name: 'Employees Table'
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