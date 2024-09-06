using ExcelQueryCLI.Common;
using ExcelQueryCLI.Models;
using Serilog;

namespace ExcelQueryCLI.Xl;

public static class ExcelQueryManager
{
  private static readonly ILogger _logger = Log.ForContext("Class", "ExcelQueryManager");

  public static void RunQuery(ExcelQueryRoot query,
                              byte parallelThreads = StaticSettings.DefaultParallelThreads) {
    _logger.Information("Processing files");
    var files = ExcelTools.GetExcelFilesList(query.Source); //must cast before looping to avoid errors in progress
    Parallel.ForEach(files,
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         var bkFilePath = string.Empty;
                         if (query.Backup) bkFilePath = ExcelTools.BackupFileToTemp(file);

                         var excelFileManager = new ExcelQueryFileManager(file, query.Sheets, query.QueryUpdate, query.QueryDelete);
                         var updatedCount = excelFileManager.Run();
                         if (!string.IsNullOrEmpty(bkFilePath)) {
                           if (updatedCount > 0)
                             ExcelTools.MoveFileToBackup(bkFilePath);
                           else
                             File.Delete(bkFilePath);
                         }
                       }
                       catch (Exception ex) {
                         _logger.Error(ex, "Exception while processing file {file}", file);
                       }
                     });

    _logger.Information("All files processed");
  }
}