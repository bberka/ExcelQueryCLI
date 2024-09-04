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

public sealed class ExcelQueryRootDelete
{
  private SheetRecord[] _sheets = [];
  private string[] _source = [];
  private DeleteQueryInformation[] _query = [];
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
      _source.Throw().IfNull(x => x).IfEmpty().IfHasNullElements();
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
      _sheets.Throw().IfNull(x => x).IfEmpty().IfHasNullElements();
    }
  }

  [YamlMember(Alias = "query")]
  [XmlElement("query")]
  [JsonProperty("query")]
  public required DeleteQueryInformation[] Query {
    get => _query;
    [MemberNotNull(nameof(_query))]
    set {
      _query = value ?? [];
      _query.Throw().IfNull(x => x).IfEmpty().IfHasNullElements();
    }
  }

  [YamlMember(Alias = "backup")]
  [XmlElement("backup")]
  [JsonProperty("backup")]
  public bool Backup { get; set; } = StaticSettings.DefaultBackup;


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

  public static ExcelQueryRootDelete ParseFile(string path, SupportedFileType fileType) {
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

  public static ExcelQueryRootDelete ParseYamlText(string yaml) {
    var deserializer = new DeserializerBuilder()
                       .WithNamingConvention(UnderscoredNamingConvention.Instance)
                       .WithEnforceRequiredMembers()
                       .Build();
    var yamlObject = deserializer.Deserialize<ExcelQueryRootDelete>(yaml);
    Log.Verbose("Parsed YAML: {@yamlObject}", yamlObject);
    yamlObject.Validate();
    Log.Verbose("Validated YAML: {@yamlObject}", yamlObject);
    return yamlObject;
  }

  public static ExcelQueryRootDelete ParseJsonText(string text) {
    var q = JsonConvert.DeserializeObject<ExcelQueryRootDelete>(text,
                                                                settings: new JsonSerializerSettings() {
                                                                  Culture = CultureInfo.InvariantCulture,
                                                                  Converters = { new Newtonsoft.Json.Converters.StringEnumConverter() }
                                                                }) ?? throw new ArgumentException("Invalid JSON");
    Log.Verbose("Parsed JSON: {@q}", q);
    q.Validate();
    Log.Verbose("Validated JSON: {@q}", q);
    return q;
  }

  public static ExcelQueryRootDelete ParseXmlText(string text) {
    var xmlSerializer = new XmlSerializer(typeof(ExcelQueryRootDelete));
    using var reader = new StringReader(text);
    xmlSerializer.UnknownAttribute += (_, args) => throw new ArgumentException("Invalid XML" + args.Attr.Name);
    xmlSerializer.UnknownElement += (_, args) => throw new ArgumentException("Invalid XML" + args.Element.Name);
    xmlSerializer.UnknownNode += (_, args) => throw new ArgumentException("Invalid XML: " + args.Name);
    var q = (ExcelQueryRootDelete?)xmlSerializer.Deserialize(reader) ?? throw new ArgumentException("Invalid XML");
    Log.Verbose("Parsed XML: {@q}", q);
    q.Validate();
    Log.Verbose("Validated XML: {@q}", q);
    return q;
  }

  public void Validate() {
    foreach (var sheet in Sheets) sheet.Validate();

    foreach (var q in Query) q.Validate(ValuesDefinitions);

    if (Sheets.Length == 0) {
      foreach (var q in Query) {
        q.Throw("Sheet must be provided").IfEmpty(x => x.Sheets);
      }
    }
  }
}