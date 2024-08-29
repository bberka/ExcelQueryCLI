using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelQueryCLI.Parsers;
using ExcelQueryCLI.Static;
using Serilog;

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


  public Tuple<bool, string> DeleteQuery(FilterQueryParser filterQuery, bool onlyFirst) {
    Log.Information("Parsed Filter Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
                    filterQuery.Column,
                    filterQuery.Operator,
                    filterQuery.Value);
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

      var filterColumnIndex = GetColumnIndex(document, worksheetPart, filterQuery.Column);
      if (filterColumnIndex == -1) {
        return new Tuple<bool, string>(false, $"Filter column {filterQuery.Column} not found.");
      }


      var deleted = 0;
      foreach (var row in sheetData.Elements<Row>()) {
        var filterCell = row.Elements<Cell>().ElementAtOrDefault(filterColumnIndex);
        if (filterCell == null) {
          Log.Verbose("DeleteQuery::Cell not found, skipping row.");
          continue;
        }

        var filterValue = GetCellValue(document, filterCell);
        var checkFilterResult = CheckFilter(filterValue, filterQuery.Value, filterQuery.Operator);
        if (checkFilterResult) {
          sheetData.RemoveChild(row);
          deleted++;
          Log.Verbose("UpdateQuery::Row deleted: {row}", row.RowIndex);
          if (onlyFirst) {
            Log.Verbose("UpdateQuery::Only deleting the first matching row, breaking out of loop.");
            break;
          }
        }
      }

      worksheetPart.Worksheet.Save();

      if (deleted > 0) {
        Log.Information("Rows deleted: {updated}", deleted);
        return new Tuple<bool, string>(true, $"Rows deleted: {deleted}");
      }

      Log.Information("No rows deleted.");
      return new Tuple<bool, string>(false, "No rows deleted.");
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
      return new Tuple<bool, string>(false, "Error deleting Excel file.");
    }
  }


  public Tuple<bool, string> UpdateQuery(FilterQueryParser filterQuery, SetQueryParser setQuery, bool onlyFirst) {
    Log.Information("Parsed Filter Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
                    filterQuery.Column,
                    filterQuery.Operator,
                    filterQuery.Value);
    Log.Information("Parsed Set Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
                    setQuery.Column,
                    setQuery.Operator,
                    setQuery.Value);


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

      var filterColumnIndex = GetColumnIndex(document, worksheetPart, filterQuery.Column);
      if (filterColumnIndex == -1) {
        return new Tuple<bool, string>(false, $"Filter column {filterQuery.Column} not found.");
      }

      var setColumnIndex = GetColumnIndex(document, worksheetPart, setQuery.Column);
      if (setColumnIndex == -1) {
        return new Tuple<bool, string>(false, $"Set column {setQuery.Column} not found.");
      }

      var updated = 0;
      foreach (var row in sheetData.Elements<Row>()) {
        var filterCell = row.Elements<Cell>().ElementAtOrDefault(filterColumnIndex);
        var setCell = row.Elements<Cell>().ElementAtOrDefault(setColumnIndex);
        if (filterCell == null || setCell == null) {
          Log.Verbose("UpdateQuery::Cell not found, skipping row.");
          continue;
        }

        var filterValue = GetCellValue(document, filterCell);
        var checkFilterResult = CheckFilter(filterValue, filterQuery.Value, filterQuery.Operator);
        if (checkFilterResult) {
          var setCellValue = GetCellValue(document, setCell);
          UpdateCellValue(setCell, setCellValue, setQuery.Value, setQuery.Operator);
          updated++;
          Log.Verbose("UpdateQuery::Row updated: {row}", row.RowIndex);
          if (onlyFirst) {
            Log.Verbose("UpdateQuery::Only updating the first matching row, breaking out of loop.");
            break;
          }
        }
      }

      worksheetPart.Worksheet.Save();

      if (updated > 0) {
        Log.Information("Rows updated: {updated}", updated);
        return new Tuple<bool, string>(true, $"Rows updated: {updated}");
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
        return double.TryParse(cellFilterValue, out var cellValue1) && double.TryParse(matchFilterValue, out var matchValue1) && cellValue1 > matchValue1;
      case FilterOperator.LESS_THAN:
        return double.TryParse(cellFilterValue, out var cellValue2) && double.TryParse(matchFilterValue, out var matchValue2) && cellValue2 < matchValue2;
      case FilterOperator.GREATER_THAN_OR_EQUAL:
        return double.TryParse(cellFilterValue, out var cellValue3) && double.TryParse(matchFilterValue, out var matchValue3) && cellValue3 >= matchValue3;
      case FilterOperator.LESS_THAN_OR_EQUAL:
        return double.TryParse(cellFilterValue, out var cellValue4) && double.TryParse(matchFilterValue, out var matchValue4) && cellValue4 <= matchValue4;
      case FilterOperator.CONTAINS:
        return cellFilterValue.Contains(matchFilterValue);
      case FilterOperator.NOT_CONTAINS:
        return !cellFilterValue.Contains(matchFilterValue);
      case FilterOperator.STARTS_WITH:
        return cellFilterValue.StartsWith(matchFilterValue);
      case FilterOperator.ENDS_WITH:
        return cellFilterValue.EndsWith(matchFilterValue);
      case FilterOperator.IN:
        return matchFilterValue.Split("|").Any(x => x.Trim() == cellFilterValue);
      case FilterOperator.BETWEEN:
        var values = matchFilterValue.Split("|");
        if (values.Length != 2) {
          return false;
        }

        return double.TryParse(cellFilterValue, out var cellValue5) &&
               double.TryParse(values[0], out var matchValue5) &&
               double.TryParse(values[1], out var matchValue6) &&
               cellValue5 >= matchValue5 &&
               cellValue5 <= matchValue6;
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
      if (!double.TryParse(oldValue, out var oldValueDouble)) {
        Log.Verbose("UpdateCellValue::Failed to parse old value: {oldValue}", oldValue);
        return;
      }

      if (!double.TryParse(newValue, out var newValueDouble)) {
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