using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelQueryCLI.Models.Update;

public sealed record ExcelUpdateQuery : IModel
{
  [YamlMember(Alias = "source")]
  public string[] Source { get; set; } = null!;

  [YamlMember(Alias = "sheets")]
  public required Dictionary<string, QuerySheetInformation> Sheets { get; set; } = null!;

  [YamlMember(Alias = "query")]
  public required UpdateQueryInformation[] Query { get; set; } = [];

  [YamlMember(Alias = "backup")]
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

  public static ExcelUpdateQuery ParseYamlFile(string path) {
    var yaml = File.ReadAllText(path);
    return ParseYamlText(yaml);
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

    foreach (var update in Query) update.Validate();
  }
}