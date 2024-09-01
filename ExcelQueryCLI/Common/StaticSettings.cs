using ExcelQueryCLI.Static;

namespace ExcelQueryCLI.Common;

public static class StaticSettings
{
  public const int DefaultHeaderRowNumber = 1;
  public const int DefaultStartRowIndex = 2;
  public const int DefaultParallelThreads = -1;
  public const bool DefaultOnlyFirst = false;

  public const bool DefaultBackup = false;
  public const char DefaultNumberStringSplitCharacter = '-';
  public const string DefaultReplaceSplitString = "|>|";

  // public const MergeOperator DefaultFilterQueryColumnMergeOperator = MergeOperator.OR;
  public static readonly string[] SupportedExtensions = [".xlsx", ".xlsm", ".xlsb", ".xls"];
}