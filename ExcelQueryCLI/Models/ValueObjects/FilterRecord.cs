using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.ValueObjects;

public record FilterRecord
{
  private string[] _values = [];
  private string _column = string.Empty;
  private CompareOperator _compareOperator;
  private string[] _valuesDefinitionKey = [];

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

  [YamlMember(Alias = "values_def_key")]
  [XmlElement("values_def_key")]
  [JsonProperty("values_def_key")]
  public string[] ValuesDefinitionKeys {
    get => _valuesDefinitionKey;
    set {
      _valuesDefinitionKey = value?.Select(x => x.Trim().Replace(" ",""))
                                  .Where(x => !string.IsNullOrEmpty(x) && !string.IsNullOrWhiteSpace(x))
                                  .Distinct()
                                  .ToArray() ?? [];
      _valuesDefinitionKey.Throw().IfHasNullElements();
    }
  }

  public void Validate(ValuesListDefinition[] valuesDefinitions) {
    foreach (var key in ValuesDefinitionKeys) {
      var valuesDefinition = valuesDefinitions.FirstOrDefault(x => x.Key == key);
      valuesDefinition.ThrowIfNull("Values definition key not found: " + key);
      var concat = Values.Concat(valuesDefinition.Values).ToArray();
      Values = concat;
    }
    switch (CompareOperator) {
      case CompareOperator.EQUALS:
        Values.Throw("Values must be provided for EQUALS operator").IfEmpty();
        break;
      case CompareOperator.NOT_EQUALS:
        Values.Throw("Values must be provided for NOT_EQUALS operator").IfEmpty();
        break;
      case CompareOperator.GREATER_THAN:
      case CompareOperator.LESS_THAN:
      case CompareOperator.GREATER_THAN_OR_EQUAL:
      case CompareOperator.LESS_THAN_OR_EQUAL:
        var beforeCount = Values.Length;
        Values = Values.Where(val => double.TryParse(val, out _)).ToArray();
        var isMatch = beforeCount == Values.Length;
        Values.Throw("Values must be provided for operators: GREATER_THAN, LESS_THAN, GREATER_THAN_OR_EQUAL, LESS_THAN_OR_EQUAL").IfEmpty().IfFalse(isMatch);
        break;
      case CompareOperator.CONTAINS:
        Values.Throw("Values must be provided for CONTAINS operator").IfEmpty();
        break;
      case CompareOperator.NOT_CONTAINS:
        Values.Throw("Values must be provided for NOT_CONTAINS operator").IfEmpty();
        break;
      case CompareOperator.STARTS_WITH:
        Values.Throw("Values must be provided for STARTS_WITH operator").IfEmpty();
        break;
      case CompareOperator.ENDS_WITH:
        Values.Throw("Values must be provided for ENDS_WITH operator").IfEmpty();
        break;
      case CompareOperator.BETWEEN:
      case CompareOperator.NOT_BETWEEN:
        var split = Values.Select(x => x.Split(StaticSettings.DefaultNumberStringSplitCharacter)).ToArray();
        split.Throw("Values must contain 2 values when using BETWEEN operator").IfFalse(x => x.Any(y => y.Length == 2));
        split.Throw("Values must contain 2 valid numbers when using BETWEEN operator").IfFalse(x => x.All(y => double.TryParse(y[0], out _) && double.TryParse(y[1], out _)));
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