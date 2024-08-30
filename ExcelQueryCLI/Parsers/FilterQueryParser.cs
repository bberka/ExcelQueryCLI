using System.Text.RegularExpressions;
using ExcelQueryCLI.Static;

/// <summary>
/// Example: "('ItemName' OR 'ItemKey') EQUALS ('MyTestValue' OR 'XSDe' OR 'DdwTest')"
/// Example: "('ItemName') EQUALS ('MyTestValue')"
/// Example: "('^Index' OR 'ItemType') EQUALS ('800001')"
/// </summary>
public sealed class FilterQueryParser
{
  private readonly string _query;

  public List<string> Columns { get; private set; }
  public List<string> Values { get; private set; }
  public FilterOperator Operator { get; private set; }

  private static readonly string EnumPattern = string.Join("|", Enum.GetNames(typeof(FilterOperator)));

  private static readonly Regex QueryRegex = new(
                                                 @"^\((?<columns>.+?)\)\s+(?<operator>" + EnumPattern + @")\s+\((?<values>.+?)\)$",
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
    var valuesRaw = match.Groups["values"].Value;

    // Split the values by " OR " and trim single quotes
    var valuesInside = valuesRaw.Split(new[] { " OR " }, StringSplitOptions.RemoveEmptyEntries);
    Values = valuesInside.Select(v => v.Trim().Trim('\'')).ToList();

    if (!Enum.TryParse<FilterOperator>(operatorStr, out var filterOperator)) {
      throw new ArgumentException($"Invalid operator: {operatorStr}");
    }

    Operator = filterOperator;

    if (Columns.Count == 0) {
      throw new ArgumentException("No valid column names provided.");
    }

    if (Values.Count == 0) {
      throw new ArgumentException("No valid values provided.");
    }
  }
}