﻿using ExcelQueryCLI.Data;
using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.ValueObjects;
using ExcelQueryCLI.Static;
using OfficeOpenXml;
using Serilog;
using Throw;

namespace ExcelQueryCLI.Xl;

public class ExcelQueryFileManager(
  string FilePath,
  SheetRecord[] Sheets,
  UpdateQueryInformation[] UpdateQueries,
  DeleteQueryInformation[] DeleteQueries)
{
  private readonly ILogger _logger = Log.ForContext("FilePath", FilePath);

  /// <summary>
  /// Runs the query on the Excel file
  /// </summary>
  /// <returns>Number of sheets updated</returns>
  public int Run() {
    var sheets = Sheets.Concat(UpdateQueries.SelectMany(x => x.Sheets))
                       .Concat(DeleteQueries.SelectMany(x => x.Sheets))
                       .DistinctBy(x => x.Name)
                       .ToList();
    sheets.Throw("Must provide sheets information")
          .IfHasNullElements()
          .IfNull(x => x)
          .IfEmpty();

    _logger.Information("Processing file");
    var excelPackage = new ExcelPackage(FilePath);
    var workbook = excelPackage.Workbook;
    var updatedSheets = 0;
    foreach (var sheet in sheets) {
      var worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheet.Name || x.Name.Equals(sheet.Name, StringComparison.OrdinalIgnoreCase));
      if (worksheet is null) {
        continue;
      }

      var rowCount = worksheet.Dimension.Rows;
      _logger.Information("Processing sheet {sheetName}", sheet.Name);
      var headers = ExcelTools.GetHeadersDictionary(worksheet, sheet.HeaderRow);
      var updatedRowCount = 0;
      var updatedCellCount = 0;
      var deletedRowCount = 0;
      var simpleData = new ExcelSimpleData(worksheet, headers);
      for (var r = sheet.StartRow; r < rowCount + 1; r++) {
        var beforeDeletedCount = deletedRowCount;
        foreach (var deleteQuery in DeleteQueries) {
          var isUpdateSheet = deleteQuery.Sheets.Any(x => x.Name == sheet.Name);
          if (!isUpdateSheet) continue;
          var resultDeleteRow = DeleteRow(simpleData, r, deleteQuery);
          var isDeleted = resultDeleteRow > 0;
          if (!isDeleted) continue;
          deletedRowCount++;
        }

        var isDeletedRow = deletedRowCount > beforeDeletedCount;
        if (isDeletedRow) continue;
        foreach (var updateQuery in UpdateQueries) {
          var isUpdateSheet = updateQuery.Sheets.Any(x => x.Name == sheet.Name);
          if (!isUpdateSheet) continue;
          var resultUpdateRow = UpdateRow(simpleData, r, updateQuery);
          var isUpdated = resultUpdateRow > 0;
          if (!isUpdated) continue;
          updatedRowCount++;
          updatedCellCount += resultUpdateRow;
        }
      }

      if (updatedRowCount <= 0 && deletedRowCount <= 0) continue;
      _logger.Information("Sheet {sheetName} UpdatedRows: {updatedRowCount} UpdatedCells: {updatedCellCount} DeletedRows: {deletedRowCount}",
                          sheet.Name,
                          updatedRowCount,
                          updatedCellCount,
                          deletedRowCount);
      updatedSheets++;
    }

    if (updatedSheets > 0) {
      excelPackage.Save();
      _logger.Information("File saved");
    }
    else {
      _logger.Information("No sheets updated");
    }

    return updatedSheets;
  }

  private static int UpdateRow(ExcelSimpleData excelSimpleData,
                               int row,
                               UpdateQueryInformation updateQueryInformation) {
    var updatedCells = 0;
    var headers = excelSimpleData.Headers;
    var worksheet = excelSimpleData.Worksheet;
    switch (updateQueryInformation.FilterMergeOperator) {
      case MergeOperator.AND when updateQueryInformation.Filters.Length == 0:
        throw new InvalidOperationException("Filters must be provided when merge operator is AND");
      case MergeOperator.AND: {
        var allMatch = ExcelTools.IsAllMatched(excelSimpleData, row, updateQueryInformation.Filters);
        if (!allMatch) return 0;
        foreach (var header in headers) {
          var cellValue = worksheet.Cells[row, header.Key + 1]?.Value?.ToString();
          foreach (var updateQuery in updateQueryInformation.Update) {
            var isUpdateCol = header.Value == updateQuery.Column || header.Value.Equals(updateQuery.Column, StringComparison.OrdinalIgnoreCase);
            if (!isUpdateCol) continue;

            var newCellValue = ExcelTools.GetNewCellValue(cellValue, updateQuery.Value, updateQuery.UpdateOperator);
            var isSameValue = cellValue == newCellValue;
            if (isSameValue) continue;

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
            var isUpdateCol = header.Value == updateQuery.Column || header.Value.Equals(updateQuery.Column, StringComparison.OrdinalIgnoreCase);
            if (!isUpdateCol) continue;

            var newCellValue = ExcelTools.GetNewCellValue(cellValue, updateQuery.Value, updateQuery.UpdateOperator);
            var isSameValue = cellValue == newCellValue;
            if (isSameValue) continue;

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

  private static int DeleteRow(ExcelSimpleData excelSimpleData, int row, DeleteQueryInformation deleteQueryInformation) {
    var updatedCells = 0;
    var worksheet = excelSimpleData.Worksheet;
    switch (deleteQueryInformation.FilterMergeOperator) {
      case MergeOperator.AND when deleteQueryInformation.Filters.Length == 0:
        throw new InvalidOperationException("Filters must be provided when merge operator is AND");
      case MergeOperator.AND: {
        var allMatch = ExcelTools.IsAllMatched(excelSimpleData, row, deleteQueryInformation.Filters);
        if (!allMatch) return 0;
        break;
      }
      case null or MergeOperator.OR: {
        var anyMatch = ExcelTools.IsAnyMatched(excelSimpleData, row, deleteQueryInformation.Filters);
        if (!anyMatch) return 0;
        break;
      }
      default:
        throw new ArgumentOutOfRangeException();
    }

    worksheet.DeleteRow(row);
    updatedCells += excelSimpleData.Headers.Count;

    return updatedCells;
  }
}