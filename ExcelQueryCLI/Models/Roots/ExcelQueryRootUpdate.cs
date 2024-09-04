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

public sealed record ExcelQueryRootUpdate
{
  private string[] _source = [];
  private SheetRecord[] _sheets;
  private UpdateQueryInformation[] _query = [];
  private ValuesListDefinition[] _valuesDefinitions = [];

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
      _source.Throw().IfEmpty().IfHasNullElements();
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
      _sheets.Throw().IfHasNullElements();
    }
  }

  [YamlMember(Alias = "query")]
  [XmlElement("query")]
  [JsonProperty("query")]
  public required UpdateQueryInformation[] Query {
    get => _query;
    [MemberNotNull(nameof(_query))]
    set {
      _query = value?.ToArray() ?? [];
      _query.Throw().IfEmpty().IfHasNullElements();
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



  public static ExcelQueryRootUpdate ParseFile(string path, SupportedFileType fileType) {
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
  
  public static ExcelQueryRootUpdate ParseYamlText(string yaml) {
    var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(UnderscoredNamingConvention.Instance)
                       .WithEnforceRequiredMembers()
                       .Build();
    var yamlObject = deserializer.Deserialize<ExcelQueryRootUpdate>(yaml);
    Log.Debug("Parsed YAML: {@YamlObject}", yamlObject);
    yamlObject.Validate();
    Log.Verbose("Validated YAML: {@YamlObject}", yamlObject);
    return yamlObject;
  }

  public static ExcelQueryRootUpdate ParseJsonText(string text) {
    var q = JsonConvert.DeserializeObject<ExcelQueryRootUpdate>(text,
                                                                settings: new JsonSerializerSettings() {
                                                                  Culture = CultureInfo.InvariantCulture,
                                                                  Converters = { new Newtonsoft.Json.Converters.StringEnumConverter() }
                                                                }) ?? throw new ArgumentException("Invalid JSON");
    Log.Debug("Parsed JSON: {@Q}", q);
    q.Validate();
    Log.Verbose("Validated JSON: {@Q}", q);
    return q;
  }

  public static ExcelQueryRootUpdate ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelQueryRootUpdate), new XmlRootAttribute("root"));
    using var reader = new StringReader(text);
    var q = (ExcelQueryRootUpdate?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
    Log.Debug("Parsed XML: {@Q}", q);
    q.Validate();
    Log.Verbose("Validated XML: {@Q}", q);
    return q;
  }

  public void Validate() {
    
    foreach (var update in Query) update.Validate(ValuesDefinitions);

    foreach (var sheet in Sheets) sheet.Validate();

    if (Sheets.Length == 0) {
      foreach (var q in Query) {
        q.Throw().IfEmpty(x => x.Sheets);
      }
    }
  }
}