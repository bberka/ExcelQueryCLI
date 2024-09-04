using System.Diagnostics.CodeAnalysis;
using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.ValueObjects;

public sealed class UpdateQueryRecord
{
  private string _column = string.Empty;
  private UpdateOperator _updateOperator;
  private string _value;

  [YamlMember(Alias = "column")]
  [XmlAttribute("column")]
  [JsonProperty("column")]
  public required string Column {
    get => _column;
    set {
      _column = value?.Trim() ?? string.Empty;
      _column.Throw("Column must be provided").IfNullOrEmpty(x => x).IfNullOrWhiteSpace(x => x);
    }
  }

  [YamlMember(Alias = "operator")]
  [XmlAttribute("operator")]
  [JsonProperty("operator")]
  public required UpdateOperator UpdateOperator {
    get => _updateOperator;
    set {
      _updateOperator = value;
      _updateOperator.ThrowIfNull("Operator must be provided");
    }
  }

  [YamlMember(Alias = "value")]
  [XmlAttribute("value")]
  [JsonProperty("value")]
  public required string Value {
    get => _value;
    [MemberNotNull(nameof(_value))]
    set {
      _value = value?.Trim() ?? string.Empty;
      _value.ThrowIfNull("Value must not be null");
    }
  }

  public void Validate() {
    var isValidNumber = double.TryParse(Value, out _);
    switch (UpdateOperator) {
      case UpdateOperator.SET:
        break;
      case UpdateOperator.MULTIPLY:
      case UpdateOperator.DIVIDE:
      case UpdateOperator.ADD:
      case UpdateOperator.SUBTRACT:
        if (!isValidNumber) throw new ArgumentException("Value must be a number for operator: SUBTRACT");
        break;
      case UpdateOperator.APPEND:
      case UpdateOperator.PREPEND:
        Value.Throw().IfNullOrEmpty(x => x).IfNullOrWhiteSpace(x => x);
        break;
      case UpdateOperator.REPLACE:
        var split = Value.Split(StaticSettings.DefaultReplaceSplitString);
        if (split.Length != 2) throw new ArgumentException("Value must contain the split string: " + StaticSettings.DefaultReplaceSplitString);
        break;
      default:
        throw new ArgumentOutOfRangeException();
    }
  }
}