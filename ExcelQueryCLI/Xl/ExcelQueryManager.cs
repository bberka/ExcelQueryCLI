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
    var files = ExcelTools.GetExcelFilesList(query.Source).Distinct().ToList(); //must cast before looping to avoid errors in progress
    Parallel.ForEach(files,
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         if (query.Backup) ExcelTools.BackupFile(file);

                         var excelFileManager = new ExcelQueryFileManager(file, query.Sheets, query.QueryUpdate, query.QueryDelete);
                         excelFileManager.Run();
                       }
                       catch (Exception ex) {
                         _logger.Error(ex, "Exception while processing file {file}", file);
                       }
                     });

    _logger.Information("All files processed");
  }
}