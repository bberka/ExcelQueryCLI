using System.Text.RegularExpressions;
using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Parsers;

/// <summary>
/// Example: "('ItemName' OR 'ItemKey') EQUALS 'MyTestValue'"
/// Example: "('ItemName') EQUALS 'MyTestValue'"
/// </summary>
public sealed class FilterQueryParser
{
  private readonly string _query;

  public List<string> Columns { get; private set; }
  public string Value { get; private set; }
  public FilterOperator Operator { get; private set; }

  private static readonly string EnumPattern = string.Join("|", Enum.GetNames(typeof(FilterOperator)));

  private static readonly Regex QueryRegex = new(
                                                 @"^\((?<columns>.+?)\)\s+(?<operator>" + EnumPattern + @")\s+'(?<value>.+?)'$",
                                                 RegexOptions.Compiled | RegexOptions.IgnoreCase);

  public FilterQueryParser(string query) {
    _query = query ?? throw new ArgumentNullException(nameof(query));

    var match = QueryRegex.Match(_query);
    if (!match.Success) {
      throw new FormatException("Query format is invalid.");
    }

    // Split the columns by " OR " and trim single quotes
    var columnsRaw = match.Groups["columns"].Value.Split(new[] { " OR " }, StringSplitOptions.RemoveEmptyEntries);
    Columns = columnsRaw.Select(c => c.Trim().Trim('\'')).ToList();

    var operatorStr = match.Groups["operator"].Value.Trim().ToUpper();
    Value = match.Groups["value"].Value.Trim('\'');

    if (!Enum.TryParse<FilterOperator>(operatorStr, out var filterOperator)) {
      throw new ArgumentException($"Invalid operator: {operatorStr}");
    }

    Operator = filterOperator;

    if (Columns.Count == 0) {
      throw new ArgumentException("No valid column names provided.");
    }
  }
}