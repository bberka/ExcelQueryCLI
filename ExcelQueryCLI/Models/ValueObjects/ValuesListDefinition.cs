using System.Xml.Serialization;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.ValueObjects;

public sealed class ValuesListDefinition
{
  private string _key = string.Empty;
  private string[] _values = [];

  [YamlMember(Alias = "key")]
  [XmlAttribute("key")]
  [JsonProperty("key")]
  public string Key {
    get => _key;
    set {
      _key = value.Trim().Replace(" ","");
      _key.Throw().IfNullOrEmpty(x => x).IfNullOrWhiteSpace(x => x);
    }
  }

  [YamlMember(Alias = "values")]
  [XmlElement("values")]
  [JsonProperty("values")]
  public string[] Values {
    get => _values;
    set {
      _values = value?.Select(x => x.Trim())
                     .Where(x => !string.IsNullOrEmpty(x) && !string.IsNullOrWhiteSpace(x))
                     .Distinct()
                     .ToArray() ?? [];
      _values.Throw().IfNull(x => x).IfHasNullElements().IfEmpty();
    }
  }

}