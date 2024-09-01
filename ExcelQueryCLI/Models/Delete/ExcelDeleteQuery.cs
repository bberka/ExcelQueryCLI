using ExcelQueryCLI.Common;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Models.Update;
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

  public static ExcelDeleteQuery ParseYamlFile(string path) {
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

    foreach (var q in Query) q.Validate();
  }
}