using System.Globalization;
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
  [JsonProperty("source")]
  public string[] Source { get; set; } = [];

  [YamlMember(Alias = "sheets")]
  [XmlElement("sheets")]
  [JsonProperty("sheets")]
  public required QuerySheetInformation[] Sheets { get; set; } = null!;

  [YamlMember(Alias = "query")]
  [XmlElement("query")]
  [JsonProperty("query")]
  public required UpdateQueryInformation[] Query { get; set; } = [];

  [YamlMember(Alias = "backup")]
  [XmlElement("backup")]
  [JsonProperty("backup")]
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
    var q = JsonConvert.DeserializeObject<ExcelUpdateQuery>(text,
                                                            settings: new JsonSerializerSettings() {
                                                              Culture = CultureInfo.InvariantCulture,
                                                              Converters =  { new Newtonsoft.Json.Converters.StringEnumConverter() }
                                                            }) ?? throw new ArgumentException("Invalid JSON");
    q.Validate();
    return q;
  }

  public static ExcelUpdateQuery ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelUpdateQuery), new XmlRootAttribute("root"));
    using var reader = new StringReader(text);
    var q = (ExcelUpdateQuery?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
    q.Validate();
    return q;
  }

  public void Validate() {
    if (Source == null) {
      throw new ArgumentException("Source must be provided");
    }

    if (Sheets == null) {
      throw new ArgumentException("Sheets must be provided");
    }

    if (Source.Length == 0)
      throw new ArgumentException("Source must be provided");

    if (Sheets.Length == 0)
      throw new ArgumentException("Sheets must be provided");

    if (Query.Length == 0)
      throw new ArgumentException("Query must be provided");

    Source = Source.Select(s => s.Trim()).ToArray();
    
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