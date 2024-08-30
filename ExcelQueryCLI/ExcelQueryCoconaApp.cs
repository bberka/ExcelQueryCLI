using Cocona;
using ExcelQueryCLI.Common;
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
    [Option("header-row-index", Description = "Header row index")]
    uint headerRowIndex = 1
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
        var reader = new ClosedXmlExcelPack(path, sheet, headerRowIndex);
        reader.UpdateQuery(fqParsed, sqParsed, onlyFirst);
      }
      catch (Exception ex) {
        Log.Error("Error updating Excel file: {Message}", ex.Message);
      }
    }
  }

 
}