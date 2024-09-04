using ExcelQueryCLI.Common;
using ExcelQueryCLI.Models.Roots;
using Serilog;

namespace ExcelQueryCLI.Xl;

public static class ExcelQueryManager
{
  public static void RunUpdateQuery(ExcelQueryRootUpdate query,
                                    byte parallelThreads = StaticSettings.DefaultParallelThreads) {
    Parallel.ForEach(ExcelTools.GetExcelFilesList(query.Source).Distinct(),
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         if (query.Backup) ExcelTools.BackupFile(file);

                         var excelFileManager = new ExcelQueryFileUpdateManager(file, query.Sheets, query.Query);
                         excelFileManager.RunUpdateQuery();
                       }
                       catch (Exception ex) {
                         Log.Error(ex, "RunUpdateQuery:Exception");
                       }
                     });
  }

  public static void RunDeleteQuery(ExcelQueryRootDelete query,
                                    byte parallelThreads = StaticSettings.DefaultParallelThreads) {
    Parallel.ForEach(ExcelTools.GetExcelFilesList(query.Source),
                     new ParallelOptions() {
                       MaxDegreeOfParallelism = parallelThreads
                     },
                     file => {
                       try {
                         if (query.Backup) ExcelTools.BackupFile(file);

                         var excelFileManager = new ExcelQueryFileDeleteManager(file, query.Sheets, query.Query);
                         excelFileManager.RunDeleteQuery();
                       }
                       catch (Exception ex) {
                         Log.Error(ex, "RunDeleteQuery:Exception");
                       }
                     });
  }
}