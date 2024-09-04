using System.Xml.Serialization;
using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Models.ValueObjects;
using ExcelQueryCLI.Static;
using Newtonsoft.Json;
using Throw;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models;

public sealed record UpdateQueryInformation : IModel
{
  private UpdateQueryRecord[] _update = [];
  private SheetRecord[] _sheets = [];
  private FilterRecord[] _filters = [];

  [YamlMember(Alias = "update")]
  [XmlElement("update")]
  [JsonProperty("update")]
  public UpdateQueryRecord[] Update {
    get => _update;
    set {
      _update = value?.Distinct().ToArray() ?? [];
      _update.Throw().IfNull(x => x).IfEmpty().IfHasNullElements();
    }
  }

  [YamlMember(Alias = "filter_merge")]
  [XmlElement("filter_merge")]
  [JsonProperty("filter_merge")]
  public MergeOperator? FilterMergeOperator { get; set; }

  [YamlMember(Alias = "filters")]
  [XmlElement("filters")]
  [JsonProperty("filters")]
  public FilterRecord[] Filters {
    get => _filters;
    set {
      _filters = value?.Distinct().ToArray() ?? [];
      _filters.Throw().IfNull(x => x).IfHasNullElements();
    }
  }

  [YamlMember(Alias = "sheets")]
  [XmlElement("sheets")]
  [JsonProperty("sheets")]
  public SheetRecord[] Sheets {
    get => _sheets;
    set {
      _sheets = value?.DistinctBy(x => x.Name).ToArray() ?? [];
      _sheets.Throw().IfNull(x => x).IfHasNullElements();
    }
  }

  public void Validate() {
    var isUpdateQueryColumnsUnique = Update.DistinctBy(x => x.Column).Count() == Update.Length;
    if (!isUpdateQueryColumnsUnique) throw new ArgumentException("Update query columns must be unique");

    foreach (var filter in Filters)
      filter.Validate();

    foreach (var query in Update)
      query.Validate();

    foreach (var sheet in Sheets) {
      sheet.Validate();
    }
    var isMultipleFilter = Filters.Length > 1;
    if (isMultipleFilter) {
      var isFilterMergeOperatorProvided = FilterMergeOperator is not null;
      if (!isFilterMergeOperatorProvided) throw new ArgumentException("Filter merge operator must be provided when filters are provided.");
    }

    if (FilterMergeOperator == MergeOperator.AND) {
      //column names must be unique
      var uniqueColumns = Filters.Select(x => x.Column).Distinct().Count();
      if (uniqueColumns != Filters.Length) throw new ArgumentException("Column names must be unique when using filter merge operator: AND");
    }
  }
}