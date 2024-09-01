using System.Numerics;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Delete;
using ExcelQueryCLI.Models.Update;
using ExcelQueryCLI.Static;
using OfficeOpenXml;
using Serilog;

namespace ExcelQueryCLI.Xl;

public class EpPlusExcelQueryManager(byte parallelThreads) : IExcelQueryManager
{
  #region UPDATE QUERY

  public void RunUpdateQuery(ExcelUpdateQuery query) {
    Parallel.ForEach(ExcelTools.GetExcelFilesList(query.Source),
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         if (query.Backup) ExcelTools.BackupFile(file);

                         internalRunUpdateQuery(file,
                                                query.Sheets.Select(x => x.Value).ToList(),
                                                query.Query);
                       }
                       catch (Exception ex) {
                         Log.Error(ex, "RunUpdateQuery:Exception");
                       }
                     });
  }

  private void internalRunUpdateQuery(string filePath,
                                      List<QuerySheetInformation> sheets,
                                      UpdateQueryInformation[] updateQueries) {
    Log.Information("Processing file: {file}", filePath);
    var excelPackage = new ExcelPackage(filePath);
    var workbook = excelPackage.Workbook;
    var updatedSheets = 0;
    foreach (var sheet in sheets) {
      var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheet.Name);
      if (worksheet is null) {
        Log.Warning("Sheet {sheetName} not found in {filePath}", sheet.Name, filePath);
        continue;
      }

      var rowCount = worksheet.Dimension.Rows;
      Log.Information("Processing sheet: {sheetName}", sheet.Name);

      Log.Verbose("Processing sheet headers {sheetName}", sheet.Name);
      var headers = ExcelTools.GetHeadersDictionary(worksheet, sheet.HeaderRow);
      Log.Verbose("Processed sheet headers {sheetName} {headerCount}", sheet.Name, headers.Count);


      var updatedRowCount = 0;
      var updatedCellCount = 0;
      var simpleData = new ExcelSimpleData(worksheet, headers);
      Log.Verbose("Processing sheet rows {sheetName}", sheet.Name);
      for (var r = sheet.StartRow; r < rowCount + 1; r++)
        foreach (var updateQuery in updateQueries) {
          var resultUpdateRow = UpdateRow(simpleData, r, updateQuery);
          var isUpdated = resultUpdateRow > 0;
          if (!isUpdated) continue;
          Log.Verbose("Row updated: {row} in file {file} {sheet}", r, filePath, sheet.Name);
          updatedRowCount++;
          updatedCellCount += resultUpdateRow;
        }

      if (updatedRowCount <= 0) continue;
      Log.Information("File {file} Sheet {sheetName} updated rows: {updatedRowCount} updated cells: {updatedCellCount}",
                      filePath,
                      sheet.Name,
                      updatedRowCount,
                      updatedCellCount);
      updatedSheets++;
    }

    if (updatedSheets > 0) {
      excelPackage.Save();
      Log.Information("File saved: {file}", filePath);
    }
    else {
      Log.Information("No sheets updated: {file}", filePath);
    }
  }

  private int UpdateRow(ExcelSimpleData excelSimpleData,
                        int row,
                        UpdateQueryInformation updateQueryInformation) {
    var updatedCells = 0;
    var headers = excelSimpleData.Headers;
    var worksheet = excelSimpleData.Worksheet;
    switch (updateQueryInformation.FilterMergeOperator) {
      case MergeOperator.AND when updateQueryInformation.Filters is null:
        throw new InvalidOperationException("Filters must be provided when merge operator is AND");
      case MergeOperator.AND: {
        var allMatch = ExcelTools.IsAllMatched(excelSimpleData, row, updateQueryInformation.Filters);
        if (!allMatch) return 0;
        foreach (var header in headers) {
          var cellValue = worksheet.Cells[row, header.Key + 1]?.Value?.ToString();
          foreach (var updateQuery in updateQueryInformation.Update) {
            var isUpdateCol = header.Value == updateQuery.Column;
            if (!isUpdateCol) continue;

            var newCellValue = ExcelTools.GetNewCellValue(cellValue, updateQuery.Value, updateQuery.UpdateOperator);
            ExcelTools.UpdateCellValue(worksheet, row, header.Key + 1, newCellValue);
            updatedCells++;
          }
        }

        break;
      }
      case null or MergeOperator.OR: {
        var anyMatch = ExcelTools.IsAnyMatched(excelSimpleData, row, updateQueryInformation.Filters);
        if (!anyMatch) return 0;
        foreach (var header in headers) {
          var cellValue = worksheet.Cells[row, header.Key + 1]?.Value?.ToString();

          foreach (var updateQuery in updateQueryInformation.Update) {
            var isUpdateCol = header.Value == updateQuery.Column;
            if (!isUpdateCol) continue;

            var newCellValue = ExcelTools.GetNewCellValue(cellValue, updateQuery.Value, updateQuery.UpdateOperator);
            ExcelTools.UpdateCellValue(worksheet, row, header.Key + 1, newCellValue);
            updatedCells++;
          }
        }

        break;
      }
      default:
        throw new ArgumentOutOfRangeException();
    }

    return updatedCells;
  }

  #endregion

  #region DELETE QUERY

  public void RunDeleteQuery(ExcelDeleteQuery query) {
    Parallel.ForEach(ExcelTools.GetExcelFilesList(query.Source),
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         if (query.Backup) ExcelTools.BackupFile(file);

                         internalRunDeleteQuery(file,
                                                query.Sheets.Select(x => x.Value).ToList(),
                                                query.Query);
                       }
                       catch (Exception ex) {
                         Log.Error(ex, "RunDeleteQuery:Exception");
                       }
                     });
  }

  private void internalRunDeleteQuery(string filePath, List<QuerySheetInformation> sheets, DeleteQueryInformation[] deleteQueries) {
    Log.Information("Processing file: {file}", filePath);
    var excelPackage = new ExcelPackage(filePath);
    var workbook = excelPackage.Workbook;
    var updatedSheets = 0;
    foreach (var sheet in sheets) {
      var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheet.Name);
      if (worksheet is null) {
        Log.Warning("Sheet {sheetName} not found in {filePath}", sheet.Name, filePath);
        continue;
      }

      var rowCount = worksheet.Dimension.Rows;
      Log.Information("Processing sheet: {sheetName}", sheet.Name);

      Log.Verbose("Processing sheet headers {sheetName}", sheet.Name);
      var headers = ExcelTools.GetHeadersDictionary(worksheet, sheet.HeaderRow);
      Log.Verbose("Processed sheet headers {sheetName} {headerCount}", sheet.Name, headers.Count);


      var updatedRowCount = 0;
      var updatedCellCount = 0;
      var simpleData = new ExcelSimpleData(worksheet, headers);
      Log.Verbose("Processing sheet rows {sheetName}", sheet.Name);
      for (var r = sheet.StartRow; r < rowCount + 1; r++)
        foreach (var updateQuery in deleteQueries) {
          var resultUpdateRow = DeleteRow(simpleData, r, updateQuery);
          var isUpdated = resultUpdateRow > 0;
          if (!isUpdated) continue;
          Log.Verbose("Row updated: {row} in file {file} {sheet}", r, filePath, sheet.Name);
          updatedRowCount++;
          updatedCellCount += resultUpdateRow;
        }

      if (updatedRowCount <= 0) continue;
      Log.Information("File {file} Sheet {sheetName} updated rows: {updatedRowCount} updated cells: {updatedCellCount}",
                      filePath,
                      sheet.Name,
                      updatedRowCount,
                      updatedCellCount);
      updatedSheets++;
    }

    if (updatedSheets > 0) {
      excelPackage.Save();
      Log.Information("File saved: {file}", filePath);
    }
    else {
      Log.Information("No sheets updated: {file}", filePath);
    }
  }

  private int DeleteRow(ExcelSimpleData excelSimpleData, int row, DeleteQueryInformation deleteQueryInformation) {
    var updatedCells = 0;
    var worksheet = excelSimpleData.Worksheet;
    switch (deleteQueryInformation.FilterMergeOperator) {
      case MergeOperator.AND when deleteQueryInformation.Filters is null:
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

  #endregion
}