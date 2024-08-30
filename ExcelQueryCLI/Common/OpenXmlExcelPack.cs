// using System.Globalization;
// using DocumentFormat.OpenXml;
// using DocumentFormat.OpenXml.Packaging;
// using DocumentFormat.OpenXml.Spreadsheet;
// using ExcelQueryCLI.Parsers;
// using ExcelQueryCLI.Static;
// using Serilog;
// using Tuple = DocumentFormat.OpenXml.Spreadsheet.Tuple;
//
// namespace ExcelQueryCLI.Common;
//
// public sealed class OpenXmlExcelPack
// {
//   private static readonly string[] SupportedExtensions = { ".xlsx", ".xlsm", ".xlsb", ".xls" };
//   public uint HeaderRowIndex { get; private set; }
//   public string SheetName { get; private set; }
//   public string FileOrDirectoryPath { get; private set; }
//   public bool IsDirectory { get; private set; }
//
//   public const int Parallelism = 3;
//
//   public OpenXmlExcelPack(string fileOrDirectoryPath, string sheetName, uint headerRowIndex) {
//     this.FileOrDirectoryPath = fileOrDirectoryPath;
//     SheetName = sheetName;
//     HeaderRowIndex = headerRowIndex;
//     if (HeaderRowIndex == 0) {
//       Log.Warning("Header row index cannot be 0, setting to 1.");
//       HeaderRowIndex = 1;
//     }
//
//     var fileExists = File.Exists(this.FileOrDirectoryPath);
//     var directoryExists = Directory.Exists(this.FileOrDirectoryPath);
//     var isPathValid = fileExists || directoryExists;
//     if (!isPathValid) {
//       Log.Error("File or directory not found: {file}", this.FileOrDirectoryPath);
//       return;
//     }
//
//     IsDirectory = directoryExists;
//     Log.Verbose("ExcelQueryCLI: {file} {sheet}", this.FileOrDirectoryPath, SheetName);
//   }
//
//
//   public void UpdateQuery(List<FilterQueryParser>? filterQueries, List<SetQueryParser> setQueries, bool onlyFirst) {
//     if (filterQueries is null) {
//       Log.Verbose("No filter query provided.");
//     }
//     else {
//       foreach (var filterQuery in filterQueries) {
//         Log.Verbose("Parsed Filter Query: Column: {Columns}, Operator: {Operator}, Value: {Values}",
//                     filterQuery.Columns,
//                     filterQuery.Operator,
//                     filterQuery.Values);
//       }
//     }
//
//     foreach (var query in setQueries) {
//       Log.Verbose("Parsed Set Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
//                   query.Column,
//                   query.Operator,
//                   query.Value);
//     }
//
//     if (IsDirectory) {
//       var excelFiles = Directory.GetFiles(FileOrDirectoryPath, "*.*", SearchOption.AllDirectories)
//                                 .Where(s => SupportedExtensions.Contains(Path.GetExtension(s)))
//                                 .ToList();
//       Parallel.ForEach(excelFiles,
//                        new ParallelOptions() {
//                          MaxDegreeOfParallelism = Parallelism,
//                        },
//                        _runUpdate);
//     }
//     else {
//       _runUpdate(FileOrDirectoryPath);
//     }
//
//     return;
//
//     void _runUpdate(string filePath) {
//       Log.Information("Processing file: {file}", filePath);
//       try {
//         using var document = SpreadsheetDocument.Open(filePath, true);
//
//         if (document.WorkbookPart is null) {
//           Log.Error("Error opening Excel file");
//           return;
//         }
//
//         if (document.WorkbookPart.Workbook.Sheets is null) {
//           Log.Error("No sheets found in Excel file");
//           return;
//         }
//
//         var worksheetPart = GetWorksheetPartByName(document.WorkbookPart, SheetName);
//         var sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
//         if (sheetData is null) {
//           Log.Error("No data found in Excel file");
//           return;
//         }
//
//         List<Tuple<List<int>, FilterQueryParser>>? filterQueryColumnIndexTuple = null;
//         if (filterQueries is not null) {
//           filterQueryColumnIndexTuple = new List<Tuple<List<int>, FilterQueryParser>>();
//           foreach (var filterQuery in filterQueries) {
//             var indexes = new List<int>();
//             foreach (var col in filterQuery.Columns) {
//               var filterColumnIndex = GetColumnIndex(document, worksheetPart, col);
//               if (filterColumnIndex == -1) {
//                 Log.Warning("Filter column {filterQuery.Column} not found.", col);
//                 continue;
//               }
//
//               indexes.Add(filterColumnIndex);
//             }
//
//             if (indexes.Count == 0) {
//               Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
//               continue;
//             }
//
//             filterQueryColumnIndexTuple.Add(new Tuple<List<int>, FilterQueryParser>(indexes, filterQuery));
//           }
//         }
//
//         Log.Information("FilterQueryColumnIndexTuple: {filterQueryColumnIndexTuple}", filterQueryColumnIndexTuple);
//
//         var setQueryColumnIndexDict = new Dictionary<int, SetQueryParser>();
//         foreach (var setQuery in setQueries) {
//           var setColumnIndex = GetColumnIndex(document, worksheetPart, setQuery.Column);
//           if (setColumnIndex == -1) {
//             Log.Warning("Set query column {Column} not found.", setQuery.Column);
//             continue;
//           }
//
//           var exists = setQueryColumnIndexDict.TryAdd(setColumnIndex, setQuery);
//           if (!exists) {
//             Log.Warning("Set query column {Column} already exists.", setQuery.Column);
//             continue;
//           }
//         }
//
//         // var setColumnIndex = GetColumnIndex(document, worksheetPart, setQuery.Column);
//         // if (setColumnIndex == -1) {
//         //   return new Tuple<bool, string>(false, $"Set column {setQuery.Column} not found.");
//         // }
//
//         var updatedCells = 0;
//         var updatedRows = 0;
//         foreach (var row in sheetData.Elements<Row>()) {
//           //each row
//           var isRowUpdated = false;
//           if (row.RowIndex is null) {
//             continue;
//           }
//
//           var isHeaderRow = row.RowIndex == HeaderRowIndex;
//           if (isHeaderRow) {
//             Log.Verbose("UpdateQuery::Skipping header row.");
//             continue;
//           }
//
//           if (filterQueryColumnIndexTuple is not null) {
//             foreach (var kpFilterQuery in filterQueryColumnIndexTuple) {
//               //each --filter-query param
//               var filterQuery = kpFilterQuery.Item2;
//               foreach (var filterColumnIndex in kpFilterQuery.Item1) {
//                 //each column in filter query param
//                 var filterCheckCell = row.Elements<Cell>().ElementAtOrDefault(filterColumnIndex);
//                 if (filterCheckCell == null) {
//                   Log.Verbose("UpdateQuery::Cell not found, skipping row.");
//                   continue;
//                 }
//
//                 var filterCheckColumnValue = GetCellValue(document, filterCheckCell);
//                 foreach (var filterValue in filterQuery.Values) {
//                   var checkFilterResult = ExcelTools.CheckIfMatchingFilter(filterCheckColumnValue, filterValue, filterQuery.Operator);
//                   if (!checkFilterResult) continue;
//                   foreach (var kpSetQuery in setQueryColumnIndexDict) {
//                     var setQuery = kpSetQuery.Value;
//                     var index = kpSetQuery.Key;
//                     var setCell = row.Elements<Cell>().ElementAtOrDefault(index);
//                     if (setCell == null) {
//                       Log.Verbose("UpdateQuery::Cell not found, skipping row.");
//                       continue;
//                     }
//
//                     var setCellValue = GetCellValue(document, setCell);
//                     UpdateCellValue(setCell, setCellValue, setQuery.Value, setQuery.Operator);
//                     updatedCells++;
//                     isRowUpdated = true;
//                     Log.Verbose("UpdateQuery::Row updated: {row}", row.RowIndex);
//                     if (onlyFirst) {
//                       Log.Verbose("UpdateQuery::Only updating the first matching row, breaking out of loop.");
//                       break;
//                     }
//                   }
//                 }
//               }
//             }
//           }
//           else {
//             //Update without filter meaning update all rows
//             foreach (var kpSetQuery in setQueryColumnIndexDict) {
//               var setQuery = kpSetQuery.Value;
//               var index = kpSetQuery.Key;
//               var setCell = row.Elements<Cell>().ElementAtOrDefault(index);
//               if (setCell == null) {
//                 Log.Verbose("UpdateQuery::Cell not found, skipping row.");
//                 continue;
//               }
//
//               var setCellValue = GetCellValue(document, setCell);
//               UpdateCellValue(setCell, setCellValue, setQuery.Value, setQuery.Operator);
//               updatedCells++;
//               isRowUpdated = true;
//               Log.Verbose("UpdateQuery::Row updated: {row}", row.RowIndex);
//               if (onlyFirst) {
//                 Log.Verbose("UpdateQuery::Only updating the first matching row, breaking out of loop.");
//                 break;
//               }
//             }
//           }
//
//           if (isRowUpdated) {
//             updatedRows++;
//           }
//         }
//
//         worksheetPart.Worksheet.Save();
//
//         if (updatedRows > 0) {
//           Log.Information("Rows updated: {updatedRows} Cells updated: {UpdatedCells}", updatedRows, updatedCells);
//           return;
//         }
//
//         Log.Information("No rows updated.");
//       }
//       catch (Exception ex) {
//         Log.Error("Error updating Excel file: {Message}", ex.Message);
//       }
//     }
//   }
//
//
//   private int GetColumnIndex(SpreadsheetDocument document, WorksheetPart worksheetPart, string columnName) {
//     var headerRowByIndex = worksheetPart.Worksheet.GetFirstChild<SheetData>()?.Elements<Row>().ElementAtOrDefault((int)(HeaderRowIndex - 1));
//     if (headerRowByIndex == null) {
//       Log.Verbose("GetColumnIndex::No header row found.");
//       return -1;
//     }
//
//     for (var i = 0; i < headerRowByIndex.Elements<Cell>().Count(); i++) {
//       var cell = headerRowByIndex.Elements<Cell>().ElementAt(i);
//       var headerName = GetCellValue(document, cell);
//       if (headerName is null) {
//         continue;
//       }
//
//       if (headerName == columnName) {
//         Log.Verbose("GetColumnIndex::Column found: {column}", columnName);
//         return i; // Return the index of the column
//       }
//     }
//
//     Log.Verbose("GetColumnIndex::Column not found: {column}", columnName);
//     return -1; // Column not found
//   }
//
//
//   private static string? GetCellValue(SpreadsheetDocument document, Cell cell) {
//     var value = cell.InnerText;
//
//     if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
//       Log.Verbose("GetCellValue::SharedString: {value}", value);
//       return document.WorkbookPart?.SharedStringTablePart?.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
//     }
//
//     Log.Verbose("GetCellValue::Value: {value}", value);
//     return value;
//   }
//
//   private static void UpdateCellValue(Cell cell, string? oldValue, string? setValue, UpdateOperator queryOperator) {
//     Log.Verbose("UpdateCellValue::Old Value: {oldValue}, Set Value: {setValue}, Operator: {queryOperator}",
//                 oldValue,
//                 setValue,
//                 queryOperator);
//
//     var newValue = ExcelTools.GetNewCellValue(oldValue, setValue, queryOperator);
//     cell.CellValue = new CellValue(newValue ?? oldValue ?? string.Empty);
//     cell.DataType = new EnumValue<CellValues>(CellValues.String);
//   }
//
//   private static WorksheetPart GetWorksheetPartByName(WorkbookPart workbookPart, string sheetName) {
//     if (workbookPart.Workbook.Sheets is null) {
//       throw new ArgumentException("No sheets found in Excel file.");
//     }
//
//     var sheet = workbookPart.Workbook.Sheets.Elements<Sheet>()
//                             .FirstOrDefault(s => s.Name == sheetName);
//
//     if (sheet == null) {
//       throw new ArgumentException($"Sheet '{sheetName}' not found.");
//     }
//
//     if (sheet.Id is null || string.IsNullOrEmpty(sheet.Id.Value)) {
//       throw new ArgumentException($"Sheet '{sheetName}' does not have an Id.");
//     }
//
//     return (WorksheetPart)workbookPart.GetPartById(sheet.Id.Value);
//   }
//
//
// }