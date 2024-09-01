using System.Text.Json.Serialization;
using System.Xml.Serialization;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.Update;

public sealed record UpdateQueryInformation : IModel
{
  [YamlMember(Alias = "update")]
  [XmlElement("update")]
  [JsonProperty("update")]
  public UpdateQuery[] Update { get; set; } = [];

  [YamlMember(Alias = "filter_merge")]
  [XmlElement("filter_merge")]
  [JsonProperty("filter_merge")]
  public MergeOperator? FilterMergeOperator { get; set; }

  [YamlMember(Alias = "filters")]
  [XmlElement("filters")]
  [JsonProperty("filters")]
  public FilterQuery[]? Filters { get; set; }

  public void Validate() {
    if (Filters is not null)
      foreach (var filter in Filters)
        filter.Validate();
    foreach (var query in Update) query.Validate();

    if (FilterMergeOperator is null && Filters is not null && Filters.Length > 1) throw new ArgumentException("Filter merge operator must be provided when filters are provided.");

    if (Filters is null && FilterMergeOperator is not null) throw new ArgumentException("Filters must be provided when filter merge operator is provided.");

    if (Filters is not null && FilterMergeOperator == MergeOperator.AND) {
      //column names must be unique
      var uniqueColumns = Filters.Select(x => x.Column).Distinct().Count();
      if (uniqueColumns != Filters.Length) throw new ArgumentException("Column names must be unique when using filter merge operator: AND");
    }

    var isUpdateQueryColumnsUnique = Update.Select(x => x.Column).Distinct().Count() == Update.Length;
    if (!isUpdateQueryColumnsUnique) throw new ArgumentException("Update query columns must be unique");
  }
}