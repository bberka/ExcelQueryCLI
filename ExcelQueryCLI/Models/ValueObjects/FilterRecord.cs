using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.ValueObjects;

public record FilterRecord : IModel
{
  private string[] _values = [];
  private string _column = string.Empty;
  private CompareOperator _compareOperator;

  [YamlMember(Alias = "column")]
  [XmlAttribute("column")]
  [JsonProperty("column")]
  public required string Column {
    get => _column;
    set {
      _column = value?.Trim() ?? string.Empty;
      _column.Throw().IfNullOrEmpty(_ => value).IfNullOrWhiteSpace(_ => value);
    }
  }

  [YamlMember(Alias = "values")]
  [XmlElement("values")]
  [JsonProperty("values")]
  public string[] Values {
    get => _values;
    set {
      _values = value?.Select(x => x.Trim())
                     .Where(x => !string.IsNullOrEmpty(x) && !string.IsNullOrWhiteSpace(x))
                     .Distinct()
                     .ToArray() ?? [];
      _values.Throw().IfNull(x => x).IfHasNullElements();
    }
  }

  [YamlMember(Alias = "compare")]
  [XmlAttribute("compare")]
  [JsonProperty("compare")]
  public required CompareOperator CompareOperator {
    get => _compareOperator;
    set {
      _compareOperator = value;
      _compareOperator.ThrowIfNull();
    }
  }

  public void Validate() {
    switch (CompareOperator) {
      case CompareOperator.EQUALS:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.NOT_EQUALS:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.GREATER_THAN:
      case CompareOperator.LESS_THAN:
      case CompareOperator.GREATER_THAN_OR_EQUAL:
      case CompareOperator.LESS_THAN_OR_EQUAL:
        var beforeCount = Values.Length;
        Values = Values.Where(val => double.TryParse(val, out _)).ToArray();
        var isMatch = beforeCount == Values.Length;
        Values.Throw().IfEmpty().IfFalse(isMatch, "Values must be numbers for operators: GREATER_THAN, LESS_THAN, GREATER_THAN_OR_EQUAL, LESS_THAN_OR_EQUAL");
        break;
      case CompareOperator.CONTAINS:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.NOT_CONTAINS:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.STARTS_WITH:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.ENDS_WITH:
        Values.Throw().IfEmpty();
        break;
      case CompareOperator.BETWEEN:
      case CompareOperator.NOT_BETWEEN:
        var split = Values.Select(x => x.Split(StaticSettings.DefaultNumberStringSplitCharacter)).ToArray();
        if (split.Any(x => x.Length != 2))
          throw new ArgumentException("Values must contain 2 values when using BETWEEN operator");
        if (split.Any(x => !double.TryParse(x[0], out _) || !double.TryParse(x[1], out _)))
          throw new ArgumentException("Values must contain 2 valid numbers when using BETWEEN operator");
        break;
      case CompareOperator.IS_NULL_OR_BLANK:
      case CompareOperator.IS_NOT_NULL_OR_BLANK:
        Values.Throw("Values must be empty when using IS_NULL_OR_BLANK or IS_NOT_NULL_OR_BLANK").IfNotEmpty();
        break;
      default:
        throw new ArgumentOutOfRangeException();
    }
  }
}