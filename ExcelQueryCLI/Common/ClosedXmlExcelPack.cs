using System.Collections;
using ClosedXML.Excel;
using ExcelQueryCLI.Parsers;
using Serilog;

namespace ExcelQueryCLI.Common;

public sealed class ClosedXmlExcelPack
{
  private static readonly string[] SupportedExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];
  public uint HeaderRowIndex { get; }
  public string SheetName { get; }
  public string FileOrDirectoryPath { get; }
  public bool IsDirectory { get; }

  public const int Parallelism = 2;

  public ClosedXmlExcelPack(string fileOrDirectoryPath, string sheetName, uint headerRowIndex) {
    this.FileOrDirectoryPath = fileOrDirectoryPath;
    SheetName = sheetName;
    HeaderRowIndex = headerRowIndex;
    if (HeaderRowIndex == 0) {
      Log.Warning("Header row index cannot be 0, setting to 1.");
      HeaderRowIndex = 1;
    }

    var fileExists = File.Exists(this.FileOrDirectoryPath);
    var directoryExists = Directory.Exists(this.FileOrDirectoryPath);
    var isPathValid = fileExists || directoryExists;
    if (!isPathValid) {
      Log.Error("File or directory not found: {file}", this.FileOrDirectoryPath);
      return;
    }

    IsDirectory = directoryExists;
  }

  public void UpdateQuery(List<FilterQueryParser>? filterQueries, List<SetQueryParser> setQueries, bool onlyFirst) {
    if (IsDirectory) {
      var excelFiles = Directory.GetFiles(FileOrDirectoryPath, "*.*", SearchOption.AllDirectories)
                                .Where(s => SupportedExtensions.Contains(Path.GetExtension(s)))
                                .ToList();
      Parallel.ForEach(excelFiles,
                       new ParallelOptions() {
                         MaxDegreeOfParallelism = Parallelism,
                       },
                       x => {
                         try {
                           _run(x);
                         }
                         catch (Exception ex) {
                           Log.Error(ex, "Exception processing file: {file}", x);
                         }
                       });
    }
    else {
      try {
        _run(FileOrDirectoryPath);
      }
      catch (Exception ex) {
        Log.Error(ex, "Exception processing file: {file}", FileOrDirectoryPath);
      }
    }

    return;


    void _run(string filePath) {
      using var workbook = new XLWorkbook(filePath);
      var hasSheet = workbook.Worksheets.Any(x => x.Name.Equals(SheetName, StringComparison.OrdinalIgnoreCase));
      if (!hasSheet) {
        Log.Warning("File does not have sheet: {sheet} {file}", SheetName, filePath);
        return;
      }
      var worksheet = workbook.Worksheet(SheetName);
      if (worksheet is null) {
        Log.Warning("File does not have sheet: {sheet} {file}", SheetName, filePath);
        return;
      }

      Log.Information("Processing file: {file}", filePath);
      var headers = GetHeaders(worksheet);
      var filterQueryColumnIndexTuple = GetFilterQueryColumnIndexTuple(filterQueries, headers);
      var setQueryColumnIndexDict = GetSetQueryColumnIndexDict(setQueries, headers);
      var updatedCells = 0;
      var updatedRows = 0;

      foreach (var row in worksheet.Rows()) {
        var isRowUpdated = false;
        var rowNumber = row.RowNumber();
        var isHeaderRow = rowNumber == HeaderRowIndex;
        if (isHeaderRow) {
          continue;
        }

        if (filterQueryColumnIndexTuple is not null) {
          //check filter then update
          foreach (var (indexes, filterQuery) in filterQueryColumnIndexTuple) {
            var isAnyMatch = CheckIfAnyMatch(indexes, filterQuery, row);
            if (!isAnyMatch) {
              continue;
            }
            
            UpdateRow(setQueryColumnIndexDict, row);
            isRowUpdated = true;
            updatedCells += 1;
          }
        }
        else {
          UpdateRow(setQueryColumnIndexDict, row);
          isRowUpdated = true;
          updatedCells += 1;
        }

        updatedRows += isRowUpdated
                         ? 1
                         : 0;
        if (!onlyFirst) continue;
        Log.Verbose("UpdateQuery::Only updating the first matching row, breaking out of loop.");
        break;
      }


      if (updatedCells > 0) {
        workbook.Save();
        Log.Information("File saved: {file} Rows updated: {updatedRows} Cells updated: {UpdatedCells}", filePath, updatedRows, updatedCells);
        return;
      }

      Log.Information("File saved: {file} no rows updated", filePath);
    }
  }

  private List<string> GetHeaders(IXLWorksheet worksheet) {
    var headerRow = worksheet.Row((int)HeaderRowIndex);
    var headers = headerRow.Cells().Select(cell => cell.Value.ToString()).ToList();
    return headers;
  }

  private bool CheckIfAnyMatch(List<int> indexes, FilterQueryParser filterQuery, IXLRow row) {
    foreach (var index in indexes) {
      var cellValue = row.Cell(index + 1).Value.ToString();
      var isMatched = ExcelTools.CheckIfMatchingFilter(cellValue, filterQuery.Values, filterQuery.Operator);
      if (!isMatched) continue;
      return true;
    }

    return false;
  }

  private void UpdateRow(Dictionary<int, SetQueryParser> setQueryColumnIndexDict, IXLRow row) {
    foreach (var (setColumnIndex, setQuery) in setQueryColumnIndexDict) {
      var cellValue = row.Cell(setColumnIndex + 1).Value.ToString();
      var newCellValue = ExcelTools.GetNewCellValue(cellValue, setQuery.Value, setQuery.Operator);
      row.Cell(setColumnIndex + 1).Value = newCellValue;
    }
  }

  private Dictionary<int, SetQueryParser> GetSetQueryColumnIndexDict(List<SetQueryParser> setQueries, List<string> headers) {
    var result = new Dictionary<int, SetQueryParser>();
    foreach (var setQuery in setQueries) {
      var setColumnIndex = headers.FindIndex(header => header.Equals(setQuery.Column, StringComparison.OrdinalIgnoreCase));
      if (setColumnIndex == -1) {
        Log.Warning("Set column {setQuery.Column} not found.", setQuery.Column);
        continue;
      }

      result[setColumnIndex] = setQuery;
    }

    return result;
  }

  private List<Tuple<List<int>, FilterQueryParser>>? GetFilterQueryColumnIndexTuple(List<FilterQueryParser>? filterQueries, List<string> headers) {
    List<Tuple<List<int>, FilterQueryParser>>? result = null;
    if (filterQueries is null) return result;
    result = new List<Tuple<List<int>, FilterQueryParser>>();
    foreach (var filterQuery in filterQueries) {
      var indexes = new List<int>();
      foreach (var col in filterQuery.Columns) {
        var filterColumnIndex = headers.FindIndex(header => header.Equals(col, StringComparison.OrdinalIgnoreCase));
        if (filterColumnIndex == -1) {
          Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
          continue;
        }

        indexes.Add(filterColumnIndex);
      }

      if (indexes.Count == 0) {
        Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
        continue;
      }

      result.Add(new Tuple<List<int>, FilterQueryParser>(indexes, filterQuery));
    }

    return result;
  }
}