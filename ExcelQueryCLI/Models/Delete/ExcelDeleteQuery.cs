using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Models.Update;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelQueryCLI.Models.Delete;

public sealed class ExcelDeleteQuery : IModel
{
  [YamlMember(Alias = "source")]
  public string[] Source { get; set; } = null!;

  [YamlMember(Alias = "sheets")]
  public required Dictionary<string, QuerySheetInformation> Sheets { get; set; } = null!;

  [YamlMember(Alias = "query")]
  public required DeleteQueryInformation[] Query { get; set; } = [];

  [YamlMember(Alias = "backup")]
  public bool Backup { get; set; } = StaticSettings.DefaultBackup;

  public static ExcelDeleteQuery ParseYamlText(string yaml) {
    var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(UnderscoredNamingConvention.Instance)
                       .WithEnforceRequiredMembers()
                       .Build();
    var yamlObject = deserializer.Deserialize<ExcelDeleteQuery>(yaml);
    yamlObject.Validate();
    return yamlObject;
  }

  public static ExcelDeleteQuery ParseFile(string path, SupportedFileType fileType) {
    var yaml = File.ReadAllText(path);
    return fileType switch {
      SupportedFileType.YAML => ParseYamlText(yaml),
      SupportedFileType.JSON => ParseJsonText(yaml),
      SupportedFileType.XML => ParseXmlText(yaml),
      _ => throw new ArgumentOutOfRangeException(nameof(fileType), fileType, null)
    };
  }

  public static ExcelDeleteQuery ParseJsonText(string text) {
    return JsonConvert.DeserializeObject<ExcelDeleteQuery>(text) ?? throw new ArgumentException("Invalid JSON");
  }

  public static ExcelDeleteQuery ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelDeleteQuery));
    using var reader = new StringReader(text);
    return (ExcelDeleteQuery?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
  }

  public void Validate() {
    if (Source.Length == 0)
      throw new ArgumentException("Source must be provided");

    if (Sheets.Count == 0)
      throw new ArgumentException("Sheets must be provided");

    if (Query.Length == 0)
      throw new ArgumentException("Query must be provided");

    var isSourceUnique = Source.Distinct().Count() == Source.Length;
    if (!isSourceUnique)
      throw new ArgumentException("Source paths must be unique");

    var isSheetNamesUnique = Sheets.Keys.Distinct().Count() == Sheets.Count;
    if (!isSheetNamesUnique)
      throw new ArgumentException("Sheet names must be unique");

    foreach (var sheet in Sheets) sheet.Value.Validate();

    foreach (var q in Query) q.Validate();
  }
}