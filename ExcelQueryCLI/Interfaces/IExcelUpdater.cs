using ExcelQueryCLI.Parsers;

namespace ExcelQueryCLI.Interfaces;

public interface IExcelUpdater
{
  public void UpdateQuery(
    string filePath,
    string sheetName,
    List<FilterQueryParser>? filterQueries,
    List<SetQueryParser> setQueries,
    bool onlyFirst,
    int headerRowNumber,
    int startRowIndex);
}