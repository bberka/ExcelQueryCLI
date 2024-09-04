using ExcelQueryCLI.Common;
using ExcelQueryCLI.Models.Roots;
using Serilog;

namespace ExcelQueryCLI.Xl;

public static class ExcelQueryManager
{
  private static readonly ILogger _logger = Log.ForContext("Class", "ExcelQueryManager");

  public static void RunQuery(ExcelQueryRoot query,
                              byte parallelThreads = StaticSettings.DefaultParallelThreads) {
    Parallel.ForEach(ExcelTools.GetExcelFilesList(query.Source).Distinct(),
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
  }
}