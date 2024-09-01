// using ExcelQueryCLI.Interfaces;
// using Serilog;
//
// namespace ExcelQueryCLI.Common;
//
// public sealed class ExcelPack
// {
//   private static readonly string[] SupportedExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];
//   public int HeaderRowNumber { get; }
//   public int ParallelThreads { get; }
//   public string SheetName { get; }
//   public int StartRowIndex { get; }
//   public string FileOrDirectoryPath { get; }
//   public bool IsDirectory { get; }
//
//   public ExcelPack(string fileOrDirectoryPath, string sheetName, int headerRowNumber, int startRowIndex, int parallelThreads) {
//     this.FileOrDirectoryPath = fileOrDirectoryPath;
//     SheetName = sheetName;
//     StartRowIndex = startRowIndex;
//     HeaderRowNumber = headerRowNumber;
//     ParallelThreads = parallelThreads;
//     if (HeaderRowNumber == 0) {
//       Log.Warning("Header row index cannot be 0, setting to 1.");
//       HeaderRowNumber = 1;
//     }
//
//     var fileExists = File.Exists(this.FileOrDirectoryPath);
//     var directoryExists = Directory.Exists(this.FileOrDirectoryPath);
//     var isPathValid = fileExists || directoryExists;
//     if (!isPathValid) {
//       Log.Error("File or directory not found: {file}", this.FileOrDirectoryPath);
//       throw new ArgumentException("File or directory not found", nameof(fileOrDirectoryPath));
//     }
//
//     IsDirectory = directoryExists;
//   }
//
//   public void UpdateQuery(IExcelUpdater excelUpdater, List<FilterQueryParser>? filterQueries, List<SetQueryParser> setQueries, bool onlyFirst) {
//     if (IsDirectory) {
//       var excelFiles = Directory.GetFiles(FileOrDirectoryPath, "*.*", SearchOption.AllDirectories)
//                                 .Where(s => SupportedExtensions.Contains(Path.GetExtension(s)))
//                                 .ToList();
//
//       if (ParallelThreads == 0) {
//         foreach (var x in excelFiles) {
//           try {
//             excelUpdater.UpdateQuery(x,
//                                      SheetName,
//                                      filterQueries,
//                                      setQueries,
//                                      onlyFirst,
//                                      HeaderRowNumber,
//                                      StartRowIndex);
//           }
//           catch (Exception ex) {
//             Log.Error(ex, "Exception processing file: {file}", x);
//           }
//         }
//       }
//       else {
//         Parallel.ForEach(excelFiles,
//                          new ParallelOptions() {
//                            MaxDegreeOfParallelism = (int)ParallelThreads,
//                          },
//                          x => {
//                            try {
//                              excelUpdater.UpdateQuery(x,
//                                                       SheetName,
//                                                       filterQueries,
//                                                       setQueries,
//                                                       onlyFirst,
//                                                       HeaderRowNumber,
//                                                       StartRowIndex);
//                            }
//                            catch (Exception ex) {
//                              Log.Error(ex, "Exception processing file: {file}", x);
//                            }
//                          });
//       }
//     }
//     else {
//       excelUpdater.UpdateQuery(FileOrDirectoryPath,
//                                SheetName,
//                                filterQueries,
//                                setQueries,
//                                onlyFirst,
//                                HeaderRowNumber,
//                                StartRowIndex);
//     }
//   }
// }

