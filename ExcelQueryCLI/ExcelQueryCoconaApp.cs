using Cocona;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Parsers;
using Serilog;

namespace ExcelQueryCLI;

public sealed class ExcelQueryCoconaApp
{
  [Command("update", Description = "Update Excel file")]
  public void Update(
    [Option("file", ['f'], Description = "Excel file path")]
    string file,
    [Option("sheet", ['s'], Description = "Sheet name")]
    string sheet,
    [Option("filter-query", Description = "Filter query string")]
    string[] filterQueryString,
    [Option("set-query", Description = "Set query string")]
    string[] setQueryString,
    [Option("only-first", Description = "Whether to update only the first matching row")]
    bool onlyFirst = false
  ) {
    Log.Information("ExcelQueryCLI.Update: {file} {sheet} {filterQuery} {setQuery}", file, sheet, filterQueryString, setQueryString);
    var fqParsed = ParseFilterQuery(filterQueryString);
    if (fqParsed is null) {
      return;
    }

    var sqParsed = ParseSetQuery(setQueryString);
    if (sqParsed is null) {
      return;
    }

    try {
      var reader = new ExcelSheetPack(file, sheet);
      var result = reader.UpdateQuery(fqParsed, sqParsed, onlyFirst);
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
    }
  }

  // [Command("delete", Description = "Delete rows from Excel file")]
  // public void Delete(
  //   [Option("file", ['f'], Description = "Excel file path")]
  //   string file,
  //   [Option("sheet", ['s'], Description = "Sheet name")]
  //   string sheet,
  //   [Option("filter-query", Description = "Filter query string")]
  //   string[] filterQueryString,
  //   [Option("only-first", Description = "Whether to update only the first matching row")]
  //   bool onlyFirst = false
  // ) {
  //   Log.Information("ExcelQueryCLI.Delete: {file} {sheet} {filterQuery}", file, sheet, filterQueryString);
  //   var fqParsed = ParseFilterQuery(filterQueryString);
  //   if (fqParsed is null) {
  //     return;
  //   }
  //
  //   try {
  //     var reader = new ExcelSheetPack(file, sheet);
  //     var result = reader.DeleteQuery(fqParsed, onlyFirst);
  //     if (result.Item1) {
  //       Log.Information("Delete successful: {Message}", result.Item2);
  //     }
  //     else {
  //       Log.Error("Delete failed: {Message}", result.Item2);
  //     }
  //   }
  //   catch (Exception ex) {
  //     Log.Error("Error updating Excel file: {Message}", ex.Message);
  //   }
  // }

  private static List<FilterQueryParser>? ParseFilterQuery(string[] filterQueryString) {
    List<FilterQueryParser> fqParsed = [];
    try {
      foreach (var query in filterQueryString) {
        var parsed = new FilterQueryParser(query);
        Log.Information("Parsed Filter Query: Column: {Column}, Operator: {Operator}, Value: {Values}",
                        parsed.Columns,
                        parsed.Operator,
                        parsed.Values);
        fqParsed.Add(parsed);
      }

      if (fqParsed.Count == 0) {
        Log.Error("No filter query provided.");
        return null;
      }
    }
    catch (FormatException ex) {
      Log.Error("Error parsing filter query: {Message}", ex.Message);
      return null;
    }
    catch (ArgumentException ex) {
      Log.Error("Invalid operator in filter query: {Message}", ex.Message);
      return null;
    }

    return fqParsed;
  }

  private static List<SetQueryParser>? ParseSetQuery(string[] setQueryString) {
    List<SetQueryParser> sqParsed = [];
    try {
      foreach (var query in setQueryString) {
        var parsed = new SetQueryParser(query);
        Log.Information("Parsed Set Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
                        parsed.Column,
                        parsed.Operator,
                        parsed.Value);
        sqParsed.Add(parsed);
      }

      var isQueryColumnNamesUnique = sqParsed.Select(x => x.Column).Distinct().Count() == sqParsed.Count;
      if (!isQueryColumnNamesUnique) {
        Log.Error("You can not set the same column multiple times for set query.");
        return null;
      }

      if (sqParsed.Count == 0) {
        Log.Error("No set query provided.");
        return null;
      }
    }
    catch (FormatException ex) {
      Log.Error("Error parsing set query: {Message}", ex.Message);
      return null;
    }
    catch (ArgumentException ex) {
      Log.Error("Invalid operator in set query: {Message}", ex.Message);
      return null;
    }

    return sqParsed;
  }
}