using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Parsers;
using OfficeOpenXml;
using Serilog;

namespace ExcelQueryCLI.Xl;

public class EpPlusExcelUpdater : IExcelUpdater
{
  public void UpdateQuery(
    string filePath,
    string sheetName,
    List<FilterQueryParser>? filterQueries,
    List<SetQueryParser> setQueries,
    bool onlyFirst,
    int headerRowNumber,
    int startRowIndex) {
    Log.Information("Processing file: {file}", filePath);
    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
    using var package = new ExcelPackage(new FileInfo(filePath));
    var worksheet = package.Workbook.Worksheets[sheetName];
    if (worksheet is null) {
      Log.Warning("File does not have sheet: {sheet} {file}", sheetName, filePath);
      return;
    }

    var rowCount = worksheet.Dimension.Rows;
    var columnCount = worksheet.Dimension.Columns;
    // var headerRow = worksheet.Cells[headerRowNumber, 1, headerRowNumber, columnCount].ToList();
    // var headerRowValues = headerRow.Select(cell => cell.Text).ToList();
    // var headerRowDict = headerRowValues.Select((value, index) => new { value, index })
    //                                    .ToDictionary(x => x.value, x => x.index);
    if (!(startRowIndex > 0)) {
      startRowIndex = headerRowNumber;
    }

    var headers = GetHeaders(worksheet, headerRowNumber);
    var filterQueryColumnIndexTuple = ExcelTools.GetFilterQueryColumnIndexTuple(filterQueries, headers);
    var setQueryColumnIndexDict = ExcelTools.GetSetQueryColumnIndexDict(setQueries, headers);
    var updatedCells = 0;
    var updatedRows = 0;

    for (var r = startRowIndex; r < rowCount + 1; r++) {
      var processedRowNumber = 0;
      var isRowUpdated = false;

      if (filterQueryColumnIndexTuple is not null) {
        //check filter then update
        foreach (var (indexes, filterQuery) in filterQueryColumnIndexTuple) {
          var isAnyMatch = CheckIfAnyMatch(worksheet, r, indexes, filterQuery);
          if (!isAnyMatch) {
            continue;
          }

          UpdateRow(worksheet, r, setQueryColumnIndexDict);
          isRowUpdated = true;
          updatedCells += 1;
        }
      }
      else {
        UpdateRow(worksheet, r, setQueryColumnIndexDict);
        isRowUpdated = true;
        updatedCells += 1;
      }

      updatedRows += isRowUpdated
                       ? 1
                       : 0;
      if (onlyFirst) break;
    }

    if (updatedCells > 0) {
      package.Save();
      Log.Information("File saved: {file} Rows updated: {updatedRows} Cells updated: {UpdatedCells}", filePath, updatedRows, updatedCells);
      return;
    }

    Log.Information("No rows updated {file}", filePath);
  }

  private void UpdateRow(ExcelWorksheet worksheet,
                         int row,
                         Dictionary<int, SetQueryParser> setQueryColumnIndexDict) {
    foreach (var (setColumnIndex, setQuery) in setQueryColumnIndexDict) {
      var cellValue = worksheet.Cells[row, setColumnIndex + 1].Value.ToString();
      var newCellValue = ExcelTools.GetNewCellValue(cellValue, setQuery.Value, setQuery.Operator);
      worksheet.Cells[row, setColumnIndex + 1].Value = newCellValue;
    }
  }

  private bool CheckIfAnyMatch(ExcelWorksheet worksheet,
                               int row,
                               List<int> colIndexes,
                               FilterQueryParser filterQuery) {
    foreach (var colIndex in colIndexes) {
      var cellValue = worksheet.Cells[row, colIndex + 1].Value.ToString();
      var isMatched = ExcelTools.CheckIfMatchingFilter(cellValue, filterQuery.Values, filterQuery.Operator);
      if (!isMatched) continue;
      return true;
    }

    return false;
  }

  private List<string> GetHeaders(ExcelWorksheet worksheet, int headerRowNumber) {
    var headerRow = worksheet.Cells[headerRowNumber, 1, headerRowNumber, worksheet.Dimension.End.Column];
    return headerRow.Select(cell => cell.Text).ToList();
  }
}