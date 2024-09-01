using Cocona;
using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Delete;
using ExcelQueryCLI.Models.Update;
using ExcelQueryCLI.Xl;
using OfficeOpenXml;
using Serilog;

namespace ExcelQueryCLI;

public sealed class ExcelQueryCoconaApp
{
  [Command("update", Description = "Update rows in Excel file")]
  public void Update(
    [Option("query", ['q'], Description = "Yaml query file path")]
    string yamlFilePath,
    [Option("parallelism", ['p'], Description = "Number of parallel threads")]
    byte parallelThreads = 1
  ) {
    if (parallelThreads < 1) {
      Log.Error("Parallel threads must be greater than or equal to 1.");
      return;
    }

    ExcelUpdateQuery q;
    try {
      q = ExcelUpdateQuery.ParseYamlFile(yamlFilePath);
    }
    catch (Exception ex) {
      Log.Error("Error parsing Yaml file: {Message}", ex.Message);
      return;
    }

    Log.Information("Processing update query");
    try {
      var manager = new EpPlusExcelQueryManager(parallelThreads);
      manager.RunUpdateQuery(q);
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
    }
  }

  [Command("delete", Description = "Delete rows in Excel file")]
  public void Delete(
    [Option("query", ['q'], Description = "Yaml query file path")]
    string yamlFilePath,
    [Option("parallelism", ['p'], Description = "Number of parallel threads")]
    byte parallelThreads = 1
  ) {
    if (parallelThreads < 1) {
      Log.Error("Parallel threads must be greater than or equal to 1.");
      return;
    }

    ExcelDeleteQuery q;
    try {
      q = ExcelDeleteQuery.ParseYamlFile(yamlFilePath);
    }
    catch (Exception ex) {
      Log.Error("Error parsing Yaml file: {Message}", ex.Message);
      return;
    }


    Log.Information("Processing delete query");
    try {
      var manager = new EpPlusExcelQueryManager(parallelThreads);
      manager.RunDeleteQuery(q);
    }
    catch (Exception ex) {
      Log.Error("Error deleting Excel file: {Message}", ex.Message);
    }
  }
}