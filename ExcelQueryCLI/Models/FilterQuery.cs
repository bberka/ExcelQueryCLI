using System.Text.Json.Serialization;
using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models;

public sealed record FilterQuery : IModel
{
  [YamlMember(Alias = "column")]
  [XmlAttribute("column")]
  [JsonProperty("column")]
  public required string Column { get; set; }

  [YamlMember(Alias = "values")]
  [XmlElement("values")]
  [JsonProperty("values")]
  public string[] Values { get; set; } = [];

  [YamlMember(Alias = "compare")]
  [XmlAttribute("compare")]
  [JsonProperty("compare")]
  public required CompareOperator CompareOperator { get; set; }

  public void Validate() {
    if (string.IsNullOrWhiteSpace(Column) || string.IsNullOrEmpty(Column) || Column is null)
      throw new ArgumentException("Column cannot be empty");


    if (Values is null)
      throw new ArgumentException("Values cannot be null");
    Values = Values.Select(x => x.Trim())
                   .Where(x => !string.IsNullOrWhiteSpace(x) && !string.IsNullOrEmpty(x))
                   .ToArray();

    switch (CompareOperator) {
      case CompareOperator.EQUALS:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.NOT_EQUALS:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.GREATER_THAN:
      case CompareOperator.LESS_THAN:
      case CompareOperator.GREATER_THAN_OR_EQUAL:
      case CompareOperator.LESS_THAN_OR_EQUAL:
        Values = Values.Where(val => double.TryParse(val, out _)).ToArray();
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.CONTAINS:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.NOT_CONTAINS:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.STARTS_WITH:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.ENDS_WITH:
        if (Values.Length == 0)
          throw new ArgumentException("Values cannot be empty");
        break;
      case CompareOperator.BETWEEN:
      case CompareOperator.NOT_BETWEEN:
        var split = Values.Select(x => x.Split(StaticSettings.DefaultNumberStringSplitCharacter)).ToArray();
        if (split.Length != 2)
          throw new ArgumentException("Values must contain 2 values when using BETWEEN operator");
        if (split.Any(x => x.Length != 2))
          throw new ArgumentException("Values must contain 2 values when using BETWEEN operator");
        if (split.Any(x => !double.TryParse(x[0], out _) || !double.TryParse(x[1], out _)))
          throw new ArgumentException("Values must contain 2 valid numbers when using BETWEEN operator");
        break;
      case CompareOperator.IS_NULL_OR_BLANK:
        if (Values.Length != 0)
          throw new ArgumentException("Values must be empty when using ");
        break;
      case CompareOperator.IS_NOT_NULL_OR_BLANK:
        if (Values.Length != 0)
          throw new ArgumentException("Values must be empty when using IS_NULL_OR_BLANK or IS_NOT_NULL_OR_BLANK");
        break;
      default:
        throw new ArgumentOutOfRangeException();
    }
  }
}