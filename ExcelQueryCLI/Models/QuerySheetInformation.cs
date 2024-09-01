using System.Text.Json.Serialization;
using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models;

public sealed record QuerySheetInformation : IModel
{
  private string _name = null!;

  [YamlMember(Alias = "name")]
  [XmlAttribute("name")]
  [JsonProperty("name")]
  public required string Name {
    get => _name;
    set => _name = value.Trim();
  }

  [YamlMember(Alias = "header_row")]
  [XmlAttribute("header_row")]
  [JsonProperty("header_row")]
  public int HeaderRow { get; set; } = StaticSettings.DefaultHeaderRowNumber;

  [YamlMember(Alias = "start_row")]
  [XmlAttribute("start_row")]
  [JsonProperty("start_row")]
  public int StartRow { get; set; } = StaticSettings.DefaultStartRowIndex;

  public void Validate() {
    if (HeaderRow < 1)
      throw new ArgumentException("HeaderRow must be greater than 0");

    if (StartRow < 1)
      throw new ArgumentException("StartRow must be greater than 0");

    if (HeaderRow >= StartRow)
      throw new ArgumentException("HeaderRow must be less than StartRow");

    if (string.IsNullOrWhiteSpace(Name) || string.IsNullOrEmpty(Name) || Name is null)
      throw new ArgumentException("Name cannot be empty");
  }
}