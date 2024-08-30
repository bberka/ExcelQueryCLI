using ExcelQueryCLI.Parsers;
using Serilog;

namespace ExcelQueryCLI;

internal static class ParamHelper
{
  internal static List<FilterQueryParser>? ParseFilterQuery(string[]? filterQueryString) {
    if (filterQueryString is null) {
      return null;
    }

    List<FilterQueryParser> fqParsed = [];
    try {
      foreach (var query in filterQueryString) {
        var parsed = new FilterQueryParser(query);
        Log.Verbose("Parsed Filter Query: Column: {Column}, Operator: {Operator}, Value: {Values}",
                    parsed.Columns,
                    parsed.Operator,
                    parsed.Values);
        fqParsed.Add(parsed);
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

  
  internal static List<SetQueryParser>? ParseSetQuery(string[] setQueryString) {
    List<SetQueryParser> sqParsed = [];
    try {
      foreach (var query in setQueryString) {
        var parsed = new SetQueryParser(query);
        Log.Verbose("Parsed Set Query: Column: {Column}, Operator: {Operator}, Value: {Value}",
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