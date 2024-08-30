using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelQueryCLI.Parsers;
using ExcelQueryCLI.Static;
using Serilog;
using Tuple = DocumentFormat.OpenXml.Spreadsheet.Tuple;

namespace ExcelQueryCLI.Common;

public sealed class ExcelSheetPack
{
  private readonly string _filePath;
  private readonly string _sheetName;

  public ExcelSheetPack(string filePath, string sheetName) {
    _filePath = filePath;
    _sheetName = sheetName;

    if (!File.Exists(_filePath)) {
      Log.Error("File not found: {file}", _filePath);
      return;
    }

    Log.Information("ExcelQueryCLI: {file} {sheet}", _filePath, _sheetName);
  }


  // public Tuple<bool, string> DeleteQuery(List<FilterQueryParser> filterQueries, bool onlyFirst) {
  //   foreach (var filterQuery in filterQueries) {
  //     Log.Information("Parsed Filter Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
  //                     filterQuery.Column,
  //                     filterQuery.Operator,
  //                     filterQuery.Value);
  //   }
  //
  //
  //   try {
  //     using var document = SpreadsheetDocument.Open(_filePath, true);
  //
  //     if (document.WorkbookPart is null) {
  //       return new Tuple<bool, string>(false, "Error opening Excel file");
  //     }
  //
  //     if (document.WorkbookPart.Workbook.Sheets is null) {
  //       return new Tuple<bool, string>(false, "No sheets found in Excel file");
  //     }
  //
  //     var worksheetPart = GetWorksheetPartByName(document.WorkbookPart, _sheetName);
  //     var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
  //     if (sheetData is null) {
  //       return new Tuple<bool, string>(false, "No data found in Excel file");
  //     }
  //
  //
  //     var filterQueryColumnIndexTuple = new List<Tuple<int, FilterQueryParser>>();
  //     foreach (var filterQuery in filterQueries) {
  //       var filterColumnIndex = GetColumnIndex(document, worksheetPart, filterQuery.Column);
  //       if (filterColumnIndex == -1) {
  //         Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Column);
  //         continue;
  //       }
  //
  //       filterQueryColumnIndexTuple.Add(new Tuple<int, FilterQueryParser>(filterColumnIndex, filterQuery));
  //     }
  //
  //
  //     var deleted = 0;
  //     var rows = sheetData.Elements<Row>().ToList();
  //     foreach (var row in rows) {
  //       foreach (var kpFilterQuery in filterQueryColumnIndexTuple) {
  //         var filterQuery = kpFilterQuery.Item2;
  //         var filterColumnIndex = kpFilterQuery.Item1;
  //         var filterCell = row.Elements<Cell>().ElementAtOrDefault(filterColumnIndex);
  //         if (filterCell == null) {
  //           Log.Verbose("UpdateQuery::Cell not found, skipping row.");
  //           continue;
  //         }
  //
  //         var filterValue = GetCellValue(document, filterCell);
  //         var checkFilterResult = CheckFilter(filterValue, filterQuery.Value, filterQuery.Operator);
  //         if (!checkFilterResult) continue;
  //         var rowIndex = row.RowIndex;
  //         if (rowIndex is null) {
  //           continue;
  //         }
  //
  //         DeleteRow(worksheetPart, sheetData, rowIndex);
  //         deleted++;
  //         Log.Verbose("UpdateQuery::Row deleted: {row}", row.RowIndex);
  //         if (onlyFirst) {
  //           Log.Verbose("UpdateQuery::Only deleting the first matching row, breaking out of loop.");
  //           break;
  //         }
  //       }
  //     }
  //
  //     worksheetPart.Worksheet.Save();
  //
  //     if (deleted > 0) {
  //       Log.Information("Rows deleted: {updated}", deleted);
  //       return new Tuple<bool, string>(true, $"Rows deleted: {deleted}");
  //     }
  //
  //     Log.Information("No rows deleted.");
  //     return new Tuple<bool, string>(false, "No rows deleted.");
  //   }
  //   catch (Exception ex) {
  //     Log.Error("Error updating Excel file: {Message}", ex.Message);
  //     return new Tuple<bool, string>(false, "Error deleting Excel file.");
  //   }
  // }
  //
  // public void DeleteRow(WorksheetPart worksheetPart, SheetData sheetData, uint rowIndex) {
  //   // Find the row to delete
  //   var rowToDelete = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex.HasValue && r.RowIndex.Value == rowIndex);
  //   if (rowToDelete == null) return;
  //
  //   // Remove the row
  //   sheetData.RemoveChild(rowToDelete);
  //
  //   // Update the RowIndex of the rows below
  //   foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex.HasValue && r.RowIndex.Value > rowIndex)) {
  //     row.RowIndex--; // Decrement the row index
  //   }
  //
  //   // Update cell references for cells in rows below the deleted row
  //   foreach (var row in sheetData.Elements<Row>().Where(r => r.RowIndex.HasValue && r.RowIndex.Value > rowIndex)) {
  //     foreach (var cell in row.Elements<Cell>()) {
  //       // Update the cell reference
  //       cell.CellReference = UpdateCellReference(cell.CellReference, -1);
  //     }
  //   }
  //
  //   // Save changes
  //   worksheetPart.Worksheet.Save();
  // }
  // private string UpdateCellReference(string cellReference, int rowOffset) {
  //   // Extract the column part and row part from the cell reference
  //   var columnPart = new string(cellReference.Where(char.IsLetter).ToArray());
  //   var rowPart = new string(cellReference.Where(char.IsDigit).ToArray());
  //
  //   // Parse the row number and apply the offset
  //   var rowIndex = uint.Parse(rowPart);
  //   var updatedRowIndex = rowIndex + (uint)rowOffset;
  //
  //   return columnPart + updatedRowIndex; // Return the updated cell reference
  // }

  public Tuple<bool, string> UpdateQuery(List<FilterQueryParser> filterQueries, List<SetQueryParser> setQueries, bool onlyFirst) {
    foreach (var filterQuery in filterQueries) {
      Log.Information("Parsed Filter Query: Column: {Columns}, Operator: {Operator}, Value: {Values}",
                      filterQuery.Columns,
                      filterQuery.Operator,
                      filterQuery.Values);
    }

    foreach (var query in setQueries) {
      Log.Information("Parsed Set Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
                      query.Column,
                      query.Operator,
                      query.Value);
    }

    try {
      using var document = SpreadsheetDocument.Open(_filePath, true);

      if (document.WorkbookPart is null) {
        return new Tuple<bool, string>(false, "Error opening Excel file");
      }

      if (document.WorkbookPart.Workbook.Sheets is null) {
        return new Tuple<bool, string>(false, "No sheets found in Excel file");
      }

      var worksheetPart = GetWorksheetPartByName(document.WorkbookPart, _sheetName);
      var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
      if (sheetData is null) {
        return new Tuple<bool, string>(false, "No data found in Excel file");
      }

      var filterQueryColumnIndexTuple = new List<Tuple<List<int>, FilterQueryParser>>();
      foreach (var filterQuery in filterQueries) {
        var indexes = new List<int>();
        foreach (var col in filterQuery.Columns) {
          var filterColumnIndex = GetColumnIndex(document, worksheetPart, col);
          if (filterColumnIndex == -1) {
            Log.Warning("Filter column {filterQuery.Column} not found.", col);
            continue;
          }
          
          indexes.Add(filterColumnIndex);
        }

        if (indexes.Count == 0) {
          Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
          continue;
        }

        filterQueryColumnIndexTuple.Add(new Tuple<List<int>, FilterQueryParser>(indexes, filterQuery));
      }

      var setQueryColumnIndexDict = new Dictionary<int, SetQueryParser>();
      foreach (var setQuery in setQueries) {
        var setColumnIndex = GetColumnIndex(document, worksheetPart, setQuery.Column);
        if (setColumnIndex == -1) {
          Log.Warning("Set column {setQuery.Column} not found.", setQuery.Column);
          continue;
        }

        var exists = setQueryColumnIndexDict.TryAdd(setColumnIndex, setQuery);
        if (!exists) {
          Log.Warning("Set column {setQuery.Column} already exists.", setQuery.Column);
          continue;
        }
      }

      // var setColumnIndex = GetColumnIndex(document, worksheetPart, setQuery.Column);
      // if (setColumnIndex == -1) {
      //   return new Tuple<bool, string>(false, $"Set column {setQuery.Column} not found.");
      // }

      var updatedCells = 0;
      var updatedRows = 0;
      foreach (var row in sheetData.Elements<Row>()) {//each row
        var isRowUpdated = false;
        foreach (var kpFilterQuery in filterQueryColumnIndexTuple) { //each --filter-query param
          var filterQuery = kpFilterQuery.Item2;
          foreach (var filterColumnIndex in kpFilterQuery.Item1) { //each column in filter query param
            var filterCheckCell = row.Elements<Cell>().ElementAtOrDefault(filterColumnIndex);
            if (filterCheckCell == null) {
              Log.Verbose("UpdateQuery::Cell not found, skipping row.");
              continue;
            }
            var filterCheckColumnValue = GetCellValue(document, filterCheckCell);
            foreach (var filterValue in filterQuery.Values) {
              var checkFilterResult = CheckFilter(filterCheckColumnValue, filterValue, filterQuery.Operator);
              if (!checkFilterResult) continue;
              foreach (var kpSetQuery in setQueryColumnIndexDict) {
                var setQuery = kpSetQuery.Value;
                var index = kpSetQuery.Key;
                var setCell = row.Elements<Cell>().ElementAtOrDefault(index);
                if (setCell == null) {
                  Log.Verbose("UpdateQuery::Cell not found, skipping row.");
                  continue;
                }

                var setCellValue = GetCellValue(document, setCell);
                UpdateCellValue(setCell, setCellValue, setQuery.Value, setQuery.Operator);
                updatedCells++;
                isRowUpdated = true;
                Log.Verbose("UpdateQuery::Row updated: {row}", row.RowIndex);
                if (onlyFirst) {
                  Log.Verbose("UpdateQuery::Only updating the first matching row, breaking out of loop.");
                  break;
                }
              }
            }

          }
        }
        if (isRowUpdated) {
          updatedRows++;
        }
      }

      worksheetPart.Worksheet.Save();

      if (updatedRows > 0) {
        Log.Information("Rows updated: {updatedRows} Cells updated: {UpdatedCells}", updatedRows, updatedCells);
        return new Tuple<bool, string>(true, $"Rows updated: {updatedRows} Cells updated: {updatedCells}");
      }

      Log.Information("No rows updated.");
      return new Tuple<bool, string>(false, "No rows updated.");
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
      return new Tuple<bool, string>(false, "Error updating Excel file.");
    }
  }

  private static bool CheckFilter(string? cellFilterValue, string matchFilterValue, FilterOperator @operator) {
    Log.Verbose("CheckFilter::Cell Value: {cellFilterValue}, Match Value: {matchFilterValue}, Operator: {Operator}",
                cellFilterValue,
                matchFilterValue,
                @operator);

    if (cellFilterValue is null) {
      return false;
    }

    switch (@operator) {
      case FilterOperator.EQUALS:
        return cellFilterValue == matchFilterValue;
      case FilterOperator.NOT_EQUALS:
        return cellFilterValue != matchFilterValue;
      case FilterOperator.GREATER_THAN:
        return double.TryParse(cellFilterValue, out var cellValue1) && double.TryParse(matchFilterValue, CultureInfo.InvariantCulture, out var matchValue1) && cellValue1 > matchValue1;
      case FilterOperator.LESS_THAN:
        return double.TryParse(cellFilterValue, out var cellValue2) && double.TryParse(matchFilterValue, CultureInfo.InvariantCulture, out var matchValue2) && cellValue2 < matchValue2;
      case FilterOperator.GREATER_THAN_OR_EQUAL:
        return double.TryParse(cellFilterValue, out var cellValue3) && double.TryParse(matchFilterValue, CultureInfo.InvariantCulture, out var matchValue3) && cellValue3 >= matchValue3;
      case FilterOperator.LESS_THAN_OR_EQUAL:
        return double.TryParse(cellFilterValue, out var cellValue4) && double.TryParse(matchFilterValue, CultureInfo.InvariantCulture, out var matchValue4) && cellValue4 <= matchValue4;
      case FilterOperator.CONTAINS:
        return cellFilterValue.Contains(matchFilterValue);
      case FilterOperator.NOT_CONTAINS:
        return !cellFilterValue.Contains(matchFilterValue);
      case FilterOperator.STARTS_WITH:
        return cellFilterValue.StartsWith(matchFilterValue);
      case FilterOperator.ENDS_WITH:
        return cellFilterValue.EndsWith(matchFilterValue);
      case FilterOperator.BETWEEN:
        var values = matchFilterValue.Split("<>");
        if (values.Length != 2) {
          return false;
        }
        return double.TryParse(cellFilterValue, CultureInfo.InvariantCulture, out var cellValue5) &&
               double.TryParse(values[0], CultureInfo.InvariantCulture, out var matchValue5) &&
               double.TryParse(values[1], CultureInfo.InvariantCulture, out var matchValue6) &&
               cellValue5 >= matchValue5 &&
               cellValue5 <= matchValue6;
      case FilterOperator.NOT_BETWEEN:
        var values2 = matchFilterValue.Split("|");
        if (values2.Length != 2) {
          return false;
        }

        return double.TryParse(cellFilterValue, CultureInfo.InvariantCulture, out var cellValue6) &&
               double.TryParse(values2[0], CultureInfo.InvariantCulture, out var matchValue7) &&
               double.TryParse(values2[1], CultureInfo.InvariantCulture, out var matchValue8) &&
               (cellValue6 < matchValue7 ||
                cellValue6 > matchValue8);
      default:
        throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null);
    }
  }


  private int GetColumnIndex(SpreadsheetDocument document, WorksheetPart worksheetPart, string columnName) {
    // Assuming the first row contains headers
    var headerRow = worksheetPart.Worksheet.GetFirstChild<SheetData>()?.Elements<Row>().FirstOrDefault();
    if (headerRow == null) {
      Log.Verbose("GetColumnIndex::No header row found.");
      return -1;
    }

    for (var i = 0; i < headerRow.Elements<Cell>().Count(); i++) {
      var cell = headerRow.Elements<Cell>().ElementAt(i);
      var headerName = GetCellValue(document, cell);
      if (headerName is null) {
        continue;
      }

      if (headerName == columnName) {
        Log.Verbose("GetColumnIndex::Column found: {column}", columnName);
        return i; // Return the index of the column
      }
    }

    Log.Verbose("GetColumnIndex::Column not found: {column}", columnName);
    return -1; // Column not found
  }


  private static string? GetCellValue(SpreadsheetDocument document, Cell cell) {
    var value = cell.InnerText;

    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
      Log.Verbose("GetCellValue::SharedString: {value}", value);
      return document.WorkbookPart?.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
    }

    Log.Verbose("GetCellValue::Value: {value}", value);
    return value;
  }

  private static void UpdateCellValue(Cell cell, string? oldValue, string? newValue, UpdateOperator queryOperator) {
    Log.Verbose("UpdateCellValue::Old Value: {oldValue}, New Value: {newValue}, Operator: {queryOperator}",
                oldValue,
                newValue,
                queryOperator);

    var isRequiredToParse = queryOperator == UpdateOperator.MULTIPLY ||
                            queryOperator == UpdateOperator.DIVIDE ||
                            queryOperator == UpdateOperator.ADD ||
                            queryOperator == UpdateOperator.SUBTRACT;

    double? parsedOldValue = null;
    double? parsedNewValue = null;
    if (isRequiredToParse) {
      if (!double.TryParse(oldValue, CultureInfo.InvariantCulture, out var oldValueDouble)) {
        Log.Verbose("UpdateCellValue::Failed to parse old value: {oldValue}", oldValue);
        return;
      }

      if (!double.TryParse(newValue, CultureInfo.InvariantCulture, out var newValueDouble)) {
        Log.Verbose("UpdateCellValue::Failed to parse new value: {newValue}", newValue);
        return;
      }

      parsedOldValue = oldValueDouble;
      parsedNewValue = newValueDouble;
    }

    switch (queryOperator) {
      case UpdateOperator.SET:
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        cell.CellValue = new CellValue(newValue ?? string.Empty);
        break;
      case UpdateOperator.MULTIPLY:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
          cell.CellValue = new CellValue((parsedOldValue * parsedNewValue).ToString() ?? string.Empty);
        }
        else {
          Log.Warning("UpdateCellValue::Failed to update cell value: {oldValue} {newValue}", oldValue, newValue);
        }

        break;
      case UpdateOperator.DIVIDE:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
          cell.CellValue = new CellValue((parsedOldValue / parsedNewValue).ToString() ?? string.Empty);
        }
        else {
          Log.Warning("UpdateCellValue::Failed to update cell value: {oldValue} {newValue}", oldValue, newValue);
        }

        break;
      case UpdateOperator.ADD:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
          cell.CellValue = new CellValue((parsedOldValue + parsedNewValue).ToString() ?? string.Empty);
        }
        else {
          Log.Warning("UpdateCellValue::Failed to update cell value: {oldValue} {newValue}", oldValue, newValue);
        }

        break;
      case UpdateOperator.SUBTRACT:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          cell.DataType = new EnumValue<CellValues>(CellValues.Number);
          cell.CellValue = new CellValue((parsedOldValue - parsedNewValue).ToString() ?? string.Empty);
        }
        else {
          Log.Warning("UpdateCellValue::Failed to update cell value: {oldValue} {newValue}", oldValue, newValue);
        }

        break;
      case UpdateOperator.APPEND:
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        cell.CellValue = new CellValue(oldValue + newValue ?? string.Empty);
        break;
      case UpdateOperator.PREPEND:
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        cell.CellValue = new CellValue(newValue + oldValue ?? string.Empty);
        break;
      case UpdateOperator.REPLACE:
        cell.DataType = new EnumValue<CellValues>(CellValues.String);
        cell.CellValue = new CellValue(oldValue?.Replace(newValue ?? string.Empty, string.Empty) ?? string.Empty);
        break;
      default:
        throw new ArgumentOutOfRangeException(nameof(queryOperator), queryOperator, null);
    }
  }

  private static WorksheetPart GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName) {
    if (workbookPart.Workbook.Sheets is null) {
      throw new ArgumentException("No sheets found in Excel file.");
    }

    var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
                            .FirstOrDefault(s => s.Name == sheetName);

    if (sheet == null) {
      throw new ArgumentException($"Sheet '{sheetName}' not found.");
    }

    if (sheet.Id is null || string.IsNullOrEmpty(sheet.Id.Value)) {
      throw new ArgumentException($"Sheet '{sheetName}' does not have an Id.");
    }

    return (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
  }
}