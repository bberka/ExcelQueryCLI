using Cocona;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Models.Roots;
using ExcelQueryCLI.Xl;
using OfficeOpenXml;
using Serilog;
using Serilog.Events;
using Serilog.Formatting.Compact;

namespace ExcelQueryCLI;

public sealed class ExcelQueryCoconaApp
{
  private void init(LogEventLevel logLevel, bool commercial, byte parallelThreads) {
    Log.Logger = new LoggerConfiguration()
                 .MinimumLevel.Is(logLevel)
                 .WriteTo.Console()
                 .WriteTo.File(new CompactJsonFormatter(), "logs/log.json", rollingInterval: RollingInterval.Hour)
                 .CreateLogger();
    Log.Information("ExcelQueryCLI v{version}", typeof(ExcelQueryCoconaApp).Assembly.GetName().Version);

    ExcelPackage.LicenseContext = commercial
                                    ? LicenseContext.Commercial
                                    : LicenseContext.NonCommercial;

    if (parallelThreads < 1) {
      Log.Error("Parallel threads must be greater than or equal to 1.");
      return;
    }
  }

  [Command("update", Description = "Update rows in Excel file")]
  public void Update(
    [Argument("query", Description = "Query file path (YAML, JSON, or XML)")]
    string filePath,
    [Option("log-level", ['l'], Description = "Log level (default: Information)")]
    LogEventLevel logLevel = LogEventLevel.Information,
    [Option("commercial", ['c'], Description = "Use commercial license (default: false)")]
    bool commercial = false,
    [Option("parallel-threads", ['p'], Description = "Number of parallel threads (default: 1)")]
    byte parallelThreads = StaticSettings.DefaultParallelThreads
  ) {
    init(logLevel, commercial, parallelThreads);
    ExcelQueryRootUpdate q;
    try {
      var fileType = ExcelTools.GetFileType(filePath);
      q = ExcelQueryRootUpdate.ParseFile(filePath, fileType);
    }
    catch (Exception ex) {
      Log.Error("Error parsing query file: {Message}", ex.Message);
      return;
    }

    Log.Information("Processing update query");
    try {
      ExcelQueryManager.RunUpdateQuery(q, parallelThreads);
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
    }
  }

  [Command("delete", Description = "Delete rows in Excel file")]
  public void Delete(
    [Argument("query", Description = "Query file path (YAML, JSON, or XML)")]
    string filePath,
    [Option("log-level", ['l'], Description = "Log level")]
    LogEventLevel logLevel = LogEventLevel.Information,
    [Option("commercial", ['c'], Description = "Use commercial license")]
    bool commercial = false,
    [Option("parallel-threads", ['p'], Description = "Number of parallel threads")]
    byte parallelThreads = StaticSettings.DefaultParallelThreads
  ) {
    init(logLevel, commercial, parallelThreads);


    ExcelQueryRootDelete q;
    try {
      var fileType = ExcelTools.GetFileType(filePath);
      q = ExcelQueryRootDelete.ParseFile(filePath, fileType);
    }
    catch (Exception ex) {
      Log.Error("Error parsing query file: {Message}", ex.Message);
      return;
    }


    Log.Information("Processing delete query");
    try {
      ExcelQueryManager.RunDeleteQuery(q, parallelThreads);
    }
    catch (Exception ex) {
      Log.Error("Error deleting Excel file: {Message}", ex.Message);
    }
  }
}