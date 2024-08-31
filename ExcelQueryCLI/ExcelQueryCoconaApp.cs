using Cocona;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Xl;
using Serilog;

namespace ExcelQueryCLI;

public sealed class ExcelQueryCoconaApp
{
  [Command("update", Description = "Update Excel file")]
  public void Update(
    [Option("file", ['f'], Description = "Excel file or directory path")]
    string[] paths,
    [Option("sheet", ['s'], Description = "Sheet name")]
    string sheet,
    [Option("filter-query", Description = "Filter query string")]
    string[]? filterQueryString,
    [Option("set-query", Description = "Set query string")]
    string[] setQueryString,
    [Option("only-first", Description = "Whether to update only the first matching row")]
    bool onlyFirst = false,
    [Option("parallelism", Description = "Number of parallel threads")]
    uint parallelThreads = 0,
    [Option("header-row-number", Description = "Header row number")]
    uint headerRowNumber = 1,
    [Option("start-row-number", Description = "Start row number")]
    uint startRowIndex = 2
  ) {
    Log.Information("ExcelQueryCLI.Update: {file} {sheet} {filterQuery} {setQuery}", paths, sheet, filterQueryString, setQueryString);
    var fqParsed = ParamHelper.ParseFilterQuery(filterQueryString);
    var sqParsed = ParamHelper.ParseSetQuery(setQueryString);
    if (sqParsed is null) {
      return;
    }

    if (paths.Length == 0) {
      Log.Error("No file or directory path provided.");
      return;
    }

    foreach (var path in paths) {
      try {
        // var reader = new OpenXmlExcelPack(path, sheet, headerRowIndex);
        var reader = new ExcelPack(path, sheet, (int)headerRowNumber, (int)startRowIndex, (int)parallelThreads);
        var updater = new EpPlusExcelUpdater();
        reader.UpdateQuery(updater, fqParsed, sqParsed, onlyFirst);
      }
      catch (Exception ex) {
        Log.Error("Error updating Excel file: {Message}", ex.Message);
      }
    }
  }
}