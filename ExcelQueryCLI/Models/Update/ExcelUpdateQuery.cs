using System.Text.Json.Serialization;
using System.Xml.Linq;
using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelQueryCLI.Models.Update;

public sealed record ExcelUpdateQuery : IModel
{
  [YamlMember(Alias = "source")]
  [XmlElement("source")]
  [JsonPropertyName("source")]
  public string[] Source { get; set; } = null!;

  [YamlMember(Alias = "sheets")]
  [XmlElement("sheets")]
  [JsonPropertyName("sheets")]
  public required QuerySheetInformation[] Sheets { get; set; } = null!;

  [YamlMember(Alias = "query")]
  [XmlElement("query")]
  [JsonPropertyName("query")]
  public required UpdateQueryInformation[] Query { get; set; } = [];

  [YamlMember(Alias = "backup")]
  [XmlElement("backup")]
  [JsonPropertyName("backup")]
  public bool Backup { get; set; } = StaticSettings.DefaultBackup;

  public static ExcelUpdateQuery ParseYamlText(string yaml) {
    var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(UnderscoredNamingConvention.Instance)
                       .WithEnforceRequiredMembers()
                       .Build();
    var yamlObject = deserializer.Deserialize<ExcelUpdateQuery>(yaml);
    yamlObject.Validate();
    return yamlObject;
  }

  public static ExcelUpdateQuery ParseFile(string path, SupportedFileType fileType) {
    var yaml = File.ReadAllText(path);
    return fileType switch {
      SupportedFileType.YAML => ParseYamlText(yaml),
      SupportedFileType.JSON => ParseJsonText(yaml),
      SupportedFileType.XML => ParseXmlText(yaml),
      _ => throw new ArgumentOutOfRangeException(nameof(fileType), fileType, null)
    };
  }

  public static ExcelUpdateQuery ParseJsonText(string text) {
    return JsonConvert.DeserializeObject<ExcelUpdateQuery>(text) ?? throw new ArgumentException("Invalid JSON");
  }

  public static ExcelUpdateQuery ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelUpdateQuery), new XmlRootAttribute("root"));
    using var reader = new StringReader(text);
    return (ExcelUpdateQuery?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
  }

  public void Validate() {
    if (Source.Length == 0)
      throw new ArgumentException("Source must be provided");

    if (Source == null) {
      throw new ArgumentException("Source must be provided");
    }
    
    if (Sheets == null) {
      throw new ArgumentException("Sheets must be provided");
    }
    
    if (Sheets.Length == 0)
      throw new ArgumentException("Sheets must be provided");

    if (Query.Length == 0)
      throw new ArgumentException("Query must be provided");

    var isSourceUnique = Source.Distinct().Count() == Source.Length;
    if (!isSourceUnique)
      throw new ArgumentException("Source paths must be unique");


    var isSheetNamesUnique = Sheets.Distinct().Count() == Sheets.Length;
    if (!isSheetNamesUnique)
      throw new ArgumentException("Sheet names must be unique");

    foreach (var sheet in Sheets) sheet.Validate();

    foreach (var update in Query) update.Validate();
  }
}