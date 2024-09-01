using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.Update;

public sealed class UpdateQuery : IModel
{
  private string _column = string.Empty;

  [YamlMember(Alias = "column")]
  public required string Column {
    get => _column;
    set => _column = value.Trim();
  }

  [YamlMember(Alias = "operator")]
  public required UpdateOperator UpdateOperator { get; set; }

  [YamlMember(Alias = "value")]
  public required string Value { get; set; }

  public void Validate() {
    if (string.IsNullOrWhiteSpace(Column))
      throw new ArgumentException("Column cannot be empty");

    if (string.IsNullOrWhiteSpace(Value))
      throw new ArgumentException("Value cannot be empty");

    var isValidNumber = double.TryParse(Value, out _);
    switch (UpdateOperator) {
      case UpdateOperator.SET:
        break;
      case UpdateOperator.MULTIPLY:
        if (!isValidNumber) throw new ArgumentException("Value must be a number for operator: MULTIPLY");

        break;
      case UpdateOperator.DIVIDE:
        if (!isValidNumber) throw new ArgumentException("Value must be a number for operator: DIVIDE");

        break;
      case UpdateOperator.ADD:
        if (!isValidNumber) throw new ArgumentException("Value must be a number for operator: ADD");

        break;
      case UpdateOperator.SUBTRACT:
        if (!isValidNumber) throw new ArgumentException("Value must be a number for operator: SUBTRACT");
        break;
      case UpdateOperator.APPEND:
        if (string.IsNullOrEmpty(Value)) throw new ArgumentException("Value cannot be empty for operator: APPEND");
        break;
      case UpdateOperator.PREPEND:
        if (string.IsNullOrEmpty(Value)) throw new ArgumentException("Value cannot be empty for operator: PREPEND");
        break;
      case UpdateOperator.REPLACE:
        if (string.IsNullOrEmpty(Value)) throw new ArgumentException("Value cannot be empty for operator: REPLACE");
        break;
      default:
        throw new ArgumentOutOfRangeException();
    }
  }
}