using ExcelQueryCLI.Data;
using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.ValueObjects;
using ExcelQueryCLI.Static;
using OfficeOpenXml;
using Serilog;
using Throw;

namespace ExcelQueryCLI.Xl;

public class ExcelQueryFileDeleteManager(
  string FilePath,
  SheetRecord[] Sheets,
  DeleteQueryInformation[] DeleteQueries)
{
  private readonly ILogger _logger = Log.ForContext("FilePath", FilePath)
                                        .ForContext("Operation", "Delete");


  public void RunDeleteQuery() {
    
    var sheets = DeleteQueries.SelectMany(x => x.Sheets).Concat(Sheets).DistinctBy(x => x.Name).ToList();
    sheets.Throw().IfEmpty();
    _logger.Information("Processing file");
    var excelPackage = new ExcelPackage(FilePath);
    var workbook = excelPackage.Workbook;
    var updatedSheets = 0;
    foreach (var sheet in sheets) {
      var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheet.Name);
      if (worksheet is null) {
        _logger.Warning("Sheet {sheetName} not found ", sheet.Name);
        continue;
      }

      var rowCount = worksheet.Dimension.Rows;
      _logger.Information("Processing sheet: {sheetName}", sheet.Name);

      _logger.Verbose("Processing sheet headers {sheetName}", sheet.Name);
      var headers = ExcelTools.GetHeadersDictionary(worksheet, sheet.HeaderRow);
      _logger.Verbose("Processed sheet headers {sheetName} {headerCount}", sheet.Name, headers.Count);


      var updatedRowCount = 0;
      var updatedCellCount = 0;
      var simpleData = new ExcelSimpleData(worksheet, headers);
      _logger.Verbose("Processing sheet rows {sheetName}", sheet.Name);
      for (var r = sheet.StartRow; r < rowCount + 1; r++)
        foreach (var updateQuery in DeleteQueries) {
          var resultUpdateRow = DeleteRow(simpleData, r, updateQuery);
          var isUpdated = resultUpdateRow > 0;
          if (!isUpdated) continue;
          _logger.Verbose("Row updated {row} in {sheet}", r, FilePath);
          updatedRowCount++;
          updatedCellCount += resultUpdateRow;
        }

      if (updatedRowCount <= 0) continue;
      _logger.Information("Sheet {sheetName} UpdatedRows: {updatedRowCount} UpdatedCells: {updatedCellCount}",
                          sheet.Name,
                          updatedRowCount,
                          updatedCellCount);
      updatedSheets++;
    }

    if (updatedSheets > 0) {
      excelPackage.Save();
      _logger.Information("File saved");
    }
    else {
      _logger.Information("No sheets updated");
    }
  }

  private int DeleteRow(ExcelSimpleData excelSimpleData, int row, DeleteQueryInformation deleteQueryInformation) {
    var updatedCells = 0;
    var worksheet = excelSimpleData.Worksheet;
    switch (deleteQueryInformation.FilterMergeOperator) {
      case MergeOperator.AND when deleteQueryInformation.Filters.Length == 0:
        throw new InvalidOperationException("Filters must be provided when merge operator is AND");
      case MergeOperator.AND: {
        var allMatch = ExcelTools.IsAllMatched(excelSimpleData, row, deleteQueryInformation.Filters);
        if (!allMatch) return 0;
        worksheet.DeleteRow(row);
        updatedCells += excelSimpleData.Headers.Count;
        break;
      }
      case null or MergeOperator.OR: {
        var anyMatch = ExcelTools.IsAnyMatched(excelSimpleData, row, deleteQueryInformation.Filters);
        if (!anyMatch) return 0;
        worksheet.DeleteRow(row);
        updatedCells += excelSimpleData.Headers.Count;
        break;
      }
      default:
        throw new ArgumentOutOfRangeException();
    }

    return updatedCells;
  }
}