using System.Text.RegularExpressions;
using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Parsers;

/// <summary>
///  Example: 'Item Identity Number' SET '12334 1XXD2'
/// </summary>
public sealed class SetQueryParser
{
  private readonly string _query;

  public string Column { get; private set; }
  public string Value { get; private set; }
  public UpdateOperator Operator { get; private set; }

  private static readonly Regex QueryRegex = new Regex(
                                                       @"^(?<column>.+?)\s+(?<operator>SET|MULTIPLY|DIVIDE|ADD|SUBTRACT)\s+'(?<value>.+?)'$",
                                                       RegexOptions.Compiled | RegexOptions.IgnoreCase);

  public SetQueryParser(string query) {
    _query = query ?? throw new ArgumentNullException(nameof(query));
    var match = QueryRegex.Match(_query);
    if (!match.Success) {
      throw new FormatException("Query format is invalid.");
    }

    Column = match.Groups["column"].Value.Trim('\'');
    var operatorStr = match.Groups["operator"].Value.Trim('\'').ToUpper();
    Value = match.Groups["value"].Value;

    if (!Enum.TryParse<UpdateOperator>(operatorStr, out var updateOperator)) {
      throw new ArgumentException($"Invalid operator: {operatorStr}");
    }

    Operator = updateOperator;

    if (string.IsNullOrEmpty(Column)) {
      throw new ArgumentException("Column name is empty.");
    }
  }
}