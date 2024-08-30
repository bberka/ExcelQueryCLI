using System.Text.RegularExpressions;
using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Parsers;

/// <summary>
/// Example: 'Item Identity Number' EQUALS '12334 1XXD2'
/// </summary>
public sealed class FilterQueryParser
{
  private readonly string _query;

  public string Column { get; private set; }
  public string Value { get; private set; }
  public FilterOperator Operator { get; private set; }

  private static readonly string EnumPattern = string.Join("|", Enum.GetNames(typeof(FilterOperator)));

  private static readonly Regex QueryRegex = new(
                                                 @"^(?<column>.+?)\s+(?<operator>" + EnumPattern + @")\s+'(?<value>.+?)'$",
                                                 RegexOptions.Compiled | RegexOptions.IgnoreCase);

  public FilterQueryParser(string query) {
    _query = query ?? throw new ArgumentNullException(nameof(query));
    //'Item Identity Number' EQUALS '12334 1XXD2'

    var match = QueryRegex.Match(_query);
    if (!match.Success) {
      throw new FormatException("Query format is invalid.");
    }

    Column = match.Groups["column"].Value.Trim('\'');
    var operatorStr = match.Groups["operator"].Value.Trim().ToUpper();
    Value = match.Groups["value"].Value.Trim('\'');

    if (!Enum.TryParse<FilterOperator>(operatorStr, out var filterOperator)) {
      throw new ArgumentException($"Invalid operator: {operatorStr}");
    }

    Operator = filterOperator;
    if (string.IsNullOrEmpty(Column)) {
      throw new ArgumentException("Column name is empty.");
    }
  }
}