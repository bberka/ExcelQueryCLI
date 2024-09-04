using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Xml.Serialization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Models.ValueObjects;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using Serilog;
using Throw;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;

namespace ExcelQueryCLI.Models.Roots;

public sealed record ExcelQueryRoot
{
  private string[] _source = [];
  private SheetRecord[] _sheets;
  private UpdateQueryInformation[] _queryUpdate = [];
  private ValuesListDefinition[] _valuesDefinitions = [];
  private DeleteQueryInformation[] _queryDelete = [];


  [YamlMember(Alias = "source")]
  [XmlElement("source")]
  [JsonProperty("source")]
  public string[] Source {
    get => _source;
    set {
      _source = value?.Select(x => x.Trim())
                     .Where(x => !string.IsNullOrEmpty(x) && !string.IsNullOrWhiteSpace(x))
                     .Distinct()
                     .ToArray() ?? [];
      _source.Throw("Source must be provided").IfNull(x => x).IfEmpty().IfHasNullElements();
    }
  }

  [YamlMember(Alias = "sheets")]
  [XmlElement("sheets")]
  [JsonProperty("sheets")]
  public required SheetRecord[] Sheets {
    get => _sheets;
    [MemberNotNull(nameof(_sheets))]
    set {
      _sheets = value?.Select(x => x with { Name = x.Name.Trim() })
                     .DistinctBy(x => x.Name)
                     .ToArray() ?? [];
      _sheets.Throw("Sheets can not contain null elements")
             .IfNull(x => x)
             .IfHasNullElements();
    }
  }

  [YamlMember(Alias = "query_update")]
  [XmlElement("query_update")]
  [JsonProperty("query_update")]
  public required UpdateQueryInformation[] QueryUpdate {
    get => _queryUpdate;
    [MemberNotNull(nameof(_queryUpdate))]
    set {
      _queryUpdate = value?.ToArray() ?? [];
      _queryUpdate.Throw("Query delete can not have null elements").IfNull(x => x).IfHasNullElements();
    }
  }

  [YamlMember(Alias = "query_delete")]
  [XmlElement("query_delete")]
  [JsonProperty("query_delete")]
  public required DeleteQueryInformation[] QueryDelete {
    get => _queryDelete;
    [MemberNotNull(nameof(_queryDelete))]
    set {
      _queryDelete = value ?? [];
      _queryDelete.Throw("Query delete can not have null elements").IfNull(x => x).IfHasNullElements();
    }
  }

  [YamlMember(Alias = "values_def")]
  [XmlElement("values_def")]
  [JsonProperty("values_def")]
  public ValuesListDefinition[] ValuesDefinitions {
    get => _valuesDefinitions;
    set {
      var isUniqueKeys = value?.DistinctBy(x => x.Key).Count() == value?.Length;
      isUniqueKeys.Throw("Values definition keys must be unique").IfFalse();
      _valuesDefinitions = value ?? [];
      _valuesDefinitions.Throw("Values Definition must not have null elements").IfHasNullElements();
    }
  }

  [YamlMember(Alias = "backup")]
  [XmlElement("backup")]
  [JsonProperty("backup")]
  public bool Backup { get; set; } = StaticSettings.DefaultBackup;


  public static ExcelQueryRoot ParseFile(string path, SupportedFileType fileType) {
    var text = File.ReadAllText(path);
    Log.Verbose("Parsing {Path} as {FileType}", path, fileType);
    Log.Debug("Text: {Text}", text);
    return fileType switch {
      SupportedFileType.YAML => ParseYamlText(text),
      SupportedFileType.JSON => ParseJsonText(text),
      SupportedFileType.XML => ParseXmlText(text),
      _ => throw new ArgumentOutOfRangeException(nameof(fileType), fileType, null)
    };
  }

  public static ExcelQueryRoot ParseYamlText(string yaml) {
    var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(UnderscoredNamingConvention.Instance)
                       .WithEnforceRequiredMembers()
                       .Build();
    var yamlObject = deserializer.Deserialize<ExcelQueryRoot>(yaml);
    Log.Debug("Parsed YAML: {@YamlObject}", yamlObject);
    yamlObject.Validate();
    Log.Verbose("Validated YAML: {@YamlObject}", yamlObject);
    return yamlObject;
  }

  public static ExcelQueryRoot ParseJsonText(string text) {
    var q = JsonConvert.DeserializeObject<ExcelQueryRoot>(text,
                                                          settings: new JsonSerializerSettings() {
                                                            Culture = CultureInfo.InvariantCulture,
                                                            Converters = { new Newtonsoft.Json.Converters.StringEnumConverter() }
                                                          }) ?? throw new ArgumentException("Invalid JSON");
    Log.Debug("Parsed JSON: {@Q}", q);
    q.Validate();
    Log.Verbose("Validated JSON: {@Q}", q);
    return q;
  }

  public static ExcelQueryRoot ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelQueryRoot), new XmlRootAttribute("root"));
    using var reader = new StringReader(text);
    var q = (ExcelQueryRoot?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
    Log.Debug("Parsed XML: {@Q}", q);
    q.Validate();
    Log.Verbose("Validated XML: {@Q}", q);
    return q;
  }

  public void Validate() {
    QueryUpdate.Throw("Update or Delete query must be provided").IfFalse(x => x.Length > 0 || QueryDelete.Length > 0);
    foreach (var r in QueryUpdate) r.Validate(ValuesDefinitions);
    foreach (var r in QueryDelete) r.Validate(ValuesDefinitions);
    foreach (var s in Sheets) s.Validate();
    if (Sheets.Length == 0) {
      foreach (var q in QueryUpdate) {
        q.Throw("Sheet must be provided").IfEmpty(x => x.Sheets);
      }

      foreach (var q in QueryDelete) {
        q.Throw("Sheet must be provided").IfEmpty(x => x.Sheets);
      }
    }
  }
}