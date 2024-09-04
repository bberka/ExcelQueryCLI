using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.ValueObjects;

public sealed record SheetRecord
{
  private string _name = null!;
  private int _startRow = StaticSettings.DefaultStartRowIndex;
  private int _headerRow = StaticSettings.DefaultHeaderRowNumber;

  [YamlMember(Alias = "name")]
  [XmlAttribute("name")]
  [JsonProperty("name")]
  public required string Name {
    get => _name;
    set {
      _name = value?.Trim() ?? string.Empty;
      _name.Throw().IfNullOrEmpty(x => x).IfNullOrWhiteSpace(x => x);
    }
  }

  [YamlMember(Alias = "header_row")]
  [XmlAttribute("header_row")]
  [JsonProperty("header_row")]
  public int HeaderRow {
    get => _headerRow;
    set {
      _headerRow = value < 1
                     ? StaticSettings.DefaultStartRowIndex
                     : value;
      _headerRow.Throw().IfTrue(x => HeaderRow >= StartRow, "HeaderRow must be smaller than StartRow");
    }
  }

  [YamlMember(Alias = "start_row")]
  [XmlAttribute("start_row")]
  [JsonProperty("start_row")]
  public int StartRow {
    get => _startRow;
    set =>
      _startRow = value < 1
                    ? StaticSettings.DefaultStartRowIndex
                    : value;
  }

  public void Validate() {

  }
}