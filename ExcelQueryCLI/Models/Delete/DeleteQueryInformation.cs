using ExcelQueryCLI.Interfaces;
using ExcelQueryCLI.Static;
using YamlDotNet.Serialization;

namespace ExcelQueryCLI.Models.Delete;

public sealed class DeleteQueryInformation : IModel
{
  [YamlMember(Alias = "filter_merge")]
  public MergeOperator? FilterMergeOperator { get; set; }

  [YamlMember(Alias = "filters")]
  public FilterQuery[]? Filters { get; set; }

  public void Validate() {
    if (Filters is null) throw new ArgumentException("Filters must be provided when using delete query.");

    if (Filters is not null)
      foreach (var filter in Filters)
        filter.Validate();

    if (FilterMergeOperator is null && Filters is not null && Filters.Length > 1) throw new ArgumentException("Filter merge operator must be provided when filters are provided.");

    if (Filters is null && FilterMergeOperator is not null) throw new ArgumentException("Filters must be provided when filter merge operator is provided.");

    if (Filters is not null && FilterMergeOperator == MergeOperator.AND) {
      //column names must be unique
      var uniqueColumns = Filters.Select(x => x.Column).Distinct().Count();
      if (uniqueColumns != Filters.Length) throw new ArgumentException("Column names must be unique when using filter merge operator: AND");
    }
  }
}