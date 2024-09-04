using ExcelQueryCLI.Models;
using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Tests;

public sealed class ExcelQueryRootTests
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
                          <query_update>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="NOT_EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;
    var excelQuery = ExcelQueryRoot.ParseXmlText(json);
    Assert.That(excelQuery.Backup, Is.True);
    Assert.That(excelQuery.Source, Has.Length.EqualTo(3));
    Assert.That(excelQuery.Source, Has.Member("ExcelFile.xlsx"));
    Assert.That(excelQuery.Source, Has.Member("ExcelFile2.xlsx"));
    Assert.That(excelQuery.Source, Has.Member("Folder\\ExcelFiles"));
    Assert.That(excelQuery.Sheets, Has.Length.EqualTo(3));
    Assert.That(excelQuery.Sheets[0].Name, Is.EqualTo("Employees Table"));
    Assert.That(excelQuery.Sheets[0].HeaderRow, Is.EqualTo(1));
    Assert.That(excelQuery.Sheets[0].StartRow, Is.EqualTo(2));
    Assert.That(excelQuery.QueryUpdate, Has.Length.EqualTo(3));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].CompareOperator, Is.EqualTo(CompareOperator.EQUALS));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Column, Is.EqualTo("NAME"));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Length.EqualTo(2));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Member("John"));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Member("Mark"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].Column, Is.EqualTo("Fullname"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.APPEND));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].Value, Is.EqualTo("John Doe"));

    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].CompareOperator, Is.EqualTo(CompareOperator.EQUALS));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Column, Is.EqualTo("NAME"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Values, Has.Length.EqualTo(1));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Values, Has.Member("John"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].CompareOperator, Is.EqualTo(CompareOperator.NOT_EQUALS));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Column, Is.EqualTo("FULLNAME"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Values, Has.Length.EqualTo(1));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Values, Has.Member("Mark"));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].Column, Is.EqualTo("Address"));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.SET));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].Value, Is.EqualTo("Turkey"));

    Assert.That(excelQuery.QueryUpdate[2].Update?[0].Column, Is.EqualTo("Salary"));
    Assert.That(excelQuery.QueryUpdate[2].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.MULTIPLY));
    Assert.That(excelQuery.QueryUpdate[2].Update?[0].Value, Is.EqualTo("1.3"));
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseXMLText_Complex_Valid2() {
    const string json = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <source>ExcelFile2.xlsx</source>
                          <source>Folder\ExcelFiles</source>
                          <backup>true</backup>
                          <query_update>
                            <sheets name="Employees Table" header_row="1" start_row="2"/>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Address Table"/>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="NOT_EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Address Table"/>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;
    var excelQuery = ExcelQueryRoot.ParseXmlText(json);
    Assert.That(excelQuery.Backup, Is.True);
    Assert.That(excelQuery.Source, Has.Length.EqualTo(3));
    Assert.That(excelQuery.Source, Has.Member("ExcelFile.xlsx"));
    Assert.That(excelQuery.Source, Has.Member("ExcelFile2.xlsx"));
    Assert.That(excelQuery.Source, Has.Member("Folder\\ExcelFiles"));
    Assert.That(excelQuery.Sheets, Has.Length.EqualTo(0));
    Assert.That(excelQuery.QueryUpdate, Has.Length.EqualTo(3));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].CompareOperator, Is.EqualTo(CompareOperator.EQUALS));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Column, Is.EqualTo("NAME"));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Length.EqualTo(2));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Member("John"));
    Assert.That(excelQuery.QueryUpdate[0].Filters?[0].Values, Has.Member("Mark"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].Column, Is.EqualTo("Fullname"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.APPEND));
    Assert.That(excelQuery.QueryUpdate[0].Update?[0].Value, Is.EqualTo("John Doe"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[1].Column, Is.EqualTo("Department"));
    Assert.That(excelQuery.QueryUpdate[0].Update?[1].UpdateOperator, Is.EqualTo(UpdateOperator.SET));
    Assert.That(excelQuery.QueryUpdate[0].Update?[1].Value, Is.EqualTo("HR"));
    Assert.That(excelQuery.QueryUpdate[0].Sheets?[0].Name, Is.EqualTo("Employees Table"));


    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].CompareOperator, Is.EqualTo(CompareOperator.EQUALS));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Column, Is.EqualTo("NAME"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Values, Has.Length.EqualTo(1));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[0].Values, Has.Member("John"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].CompareOperator, Is.EqualTo(CompareOperator.NOT_EQUALS));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Column, Is.EqualTo("FULLNAME"));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Values, Has.Length.EqualTo(1));
    Assert.That(excelQuery.QueryUpdate[1].Filters?[1].Values, Has.Member("Mark"));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].Column, Is.EqualTo("Address"));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.SET));
    Assert.That(excelQuery.QueryUpdate[1].Update?[0].Value, Is.EqualTo("Turkey"));
    Assert.That(excelQuery.QueryUpdate[1].Sheets?[0].Name, Is.EqualTo("Address Table"));


    Assert.That(excelQuery.QueryUpdate[2].Update?[0].Column, Is.EqualTo("Salary"));
    Assert.That(excelQuery.QueryUpdate[2].Update?[0].UpdateOperator, Is.EqualTo(UpdateOperator.MULTIPLY));
    Assert.That(excelQuery.QueryUpdate[2].Update?[0].Value, Is.EqualTo("1.3"));

    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleXMLText_Invalid4() {
    const string text = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <source>ExcelFile2.xlsx</source>
                          <source>Folder\ExcelFiles</source>
                          <backup>true</backup>
                          <query_update>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Salary Table"/>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;
    Assert.Throws<ArgumentException>(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }


  [Test]
  public void TestExcelQueryParseSimpleXMLText_Invalid5() {
    const string text = """
                        <root>
                          <sheets name="Salary Table"/>
                          <query_update>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Salary Table"/>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;
    Assert.Throws<InvalidOperationException>(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleXMLText_Valid() {
    const string text = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <sheets name="Salary Table"/>
                          <query_update>
                            <update column="Fullname" operator="APPEND" value="John Doe"/>
                            <update column="Department" operator="SET" value="HR"/>
                            <filters column="NAME" compare="IS_NULL_OR_BLANK">
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Salary Table"/>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="IS_NOT_NULL_OR_BLANK">
                            </filters>
                            <filters column="FULLNAME" compare="EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;

    Assert.DoesNotThrow(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleXMLText_Valid2() {
    const string text = """
                        <root>
                        <source>ExcelFile.xlsx</source>
                          <sheets name="Salary Table"/>
                          <query_update>
                            <update column="Fullname" operator="REPLACE" value="John|>|Doe"/>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Salary Table"/>
                            <update column="Address" operator="SET" value="Turkey"/>
                            <filter_merge>AND</filter_merge>
                            <filters column="NAME" compare="EQUALS">
                              <values>John</values>
                            </filters>
                            <filters column="FULLNAME" compare="EQUALS">
                              <values>Mark</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <update column="Salary" operator="MULTIPLY" value="1.3"/>
                          </query_update>
                        </root>
                        """;
    Assert.DoesNotThrow(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }


  [Test]
  public void TestExcelQueryParseSimpleXMLText_Valid3() {
    const string text = """
                        <root>
                        <source>ExcelFile.xlsx</source>
                          <sheets name="Salary Table"/>
                          <query_update>
                            <update column="Fullname" operator="REPLACE" value="John|>|"/>
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                            <filters column="Salary" compare="NOT_BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                            <filters column="Salary" compare="GREATER_THAN">
                              <values>2000</values>
                            </filters>
                            <filters column="Salary" compare="LESS_THAN">
                              <values>2000</values>
                            </filters>
                            <filters column="Salary" compare="GREATER_THAN_OR_EQUAL">
                              <values>2000</values>
                            </filters>
                            <filters column="Salary" compare="LESS_THAN_OR_EQUAL">
                              <values>2000</values>
                            </filters>
                            <filters column="Salary" compare="LESS_THAN_OR_EQUAL">
                              <values>2000</values>
                            </filters>
                          </query_update>
                        </root>
                        """;
    Assert.DoesNotThrow(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }

  [Test]
  public void TestExcelQueryParseSimpleXMLText_Invalid6() {
    const string text = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <sheets name="Salary Table"/>
                          <query_update>
                            <update column="Fullname" operator="REPLACE" value="John|>|"/>
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="GREATER_THAN">
                              <values>2000asdwqe</values>
                            </filters>
                          </query_update>
                        </root>
                        """;
    Assert.Throws<ArgumentException>(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }


  [Test]
  public void TestExcelQueryParseSimpleXMLText_Valid4() {
    const string text = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <query_update>
                            <sheets name="Salary Table" />
                            <update column="Fullname" operator="REPLACE" value="John|>|" />
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                          </query_update>
                          <query_update>
                            <sheets name="Salary Table" />
                            <update column="Fullname" operator="REPLACE" value="John|>|" />
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                          </query_update>
                        </root>
                        """;
    Assert.DoesNotThrow(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }


  [Test]
  public void TestExcelQueryParseSimpleXMLText_Valid5() {
    const string text = """
                        <root>
                          <source>ExcelFile.xlsx</source>
                          <query_update>
                            <sheets name="Salary Table" />
                            <update column="Fullname" operator="REPLACE" value="John|>|" />
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                          </query_update>
                          <query_delete>
                            <sheets name="Salary Table" />
                            <filter_merge>OR</filter_merge>
                            <filters column="Salary" compare="BETWEEN">
                              <values>2000-3000</values>
                            </filters>
                          </query_delete>
                        </root>
                        """;
    Assert.DoesNotThrow(() => { _ = ExcelQueryRoot.ParseXmlText(text); });
    Assert.Pass();
  }
}