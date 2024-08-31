using System;
using System.Text.RegularExpressions;
using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Parsers
{
  /// <summary>
  /// Example: "('Item Identity Number') SET ('12334 1XXD2')"
  /// </summary>
  public sealed class SetQueryParser
  {
    private readonly string _query;

    public string Column { get; private set; }
    public string Value { get; private set; }
    public UpdateOperator Operator { get; private set; }

    private static readonly string EnumPattern = string.Join("|", Enum.GetNames(typeof(UpdateOperator)));

    // Updated regex pattern to match the new syntax
    private static readonly Regex QueryRegex = new Regex(
                                                         @"^\((?<column>.+?)\)\s+(?<operator>" + EnumPattern + @")\s+\((?<value>.+?)\)$",
                                                         RegexOptions.Compiled | RegexOptions.IgnoreCase);

    public SetQueryParser(string query) {
      _query = query ?? throw new ArgumentNullException(nameof(query));
      var match = QueryRegex.Match(_query);
      if (!match.Success) {
        throw new FormatException("Query format is invalid.");
      }

      Column = match.Groups["column"].Value.Trim().Trim('\'');
      Value = match.Groups["value"].Value.Trim().Trim('\'');
      var operatorString = match.Groups["operator"].Value.Trim().ToUpper();

      // Parse the operator from the matched string
      if (Enum.TryParse(typeof(UpdateOperator), operatorString, out var result)) {
        Operator = (UpdateOperator)result;
      }
      else {
        throw new ArgumentException($"Invalid operator: {operatorString}");
      }

      if (string.IsNullOrEmpty(Column)) {
        throw new ArgumentException("Column name is empty.");
      }

      if (string.IsNullOrEmpty(Value)) {
        throw new ArgumentException("Value is empty.");
      }
    }
  }
}