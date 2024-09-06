using System.Globalization;
using ExcelQueryCLI.Common;
using ExcelQueryCLI.Data;
using ExcelQueryCLI.Models.ValueObjects;
using ExcelQueryCLI.Static;
using OfficeOpenXml;

namespace ExcelQueryCLI;

internal static class ExcelTools
{
  internal static bool CheckIfAnyMatch(string? checkCellValue, string[] matchValue, CompareOperator @operator) {
    return matchValue.Select(match => CheckIfMatch(checkCellValue, match, @operator)).Any(res => res);
  }

  internal static bool CheckIfMatch(string? checkCellValue, string matchValue, CompareOperator @operator) {
    if (checkCellValue is null) return false;

    switch (@operator) {
      case CompareOperator.EQUALS:
        return checkCellValue == matchValue;
      case CompareOperator.NOT_EQUALS:
        return checkCellValue != matchValue;
      case CompareOperator.GREATER_THAN:
        return double.TryParse(checkCellValue, out var cellValue1)
               && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue1)
               && cellValue1 > matchValue1;
      case CompareOperator.LESS_THAN:
        return double.TryParse(checkCellValue, out var cellValue2)
               && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue2)
               && cellValue2 < matchValue2;
      case CompareOperator.GREATER_THAN_OR_EQUAL:
        return double.TryParse(checkCellValue, out var cellValue3)
               && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue3)
               && cellValue3 >= matchValue3;
      case CompareOperator.LESS_THAN_OR_EQUAL:
        return double.TryParse(checkCellValue, out var cellValue4)
               && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue4)
               && cellValue4 <= matchValue4;
      case CompareOperator.CONTAINS:

        return checkCellValue.Contains(matchValue);
      case CompareOperator.NOT_CONTAINS:
        return !checkCellValue.Contains(matchValue);
      case CompareOperator.STARTS_WITH:
        return checkCellValue.StartsWith(matchValue);
      case CompareOperator.ENDS_WITH:
        return checkCellValue.EndsWith(matchValue);
      case CompareOperator.BETWEEN:
        var values = matchValue.Split(StaticSettings.DefaultNumberStringSplitCharacter);
        if (values.Length != 2) return false;

        return double.TryParse(checkCellValue, CultureInfo.InvariantCulture, out var cellValue5) &&
               double.TryParse(values[0], CultureInfo.InvariantCulture, out var matchValue5) &&
               double.TryParse(values[1], CultureInfo.InvariantCulture, out var matchValue6) &&
               cellValue5 >= matchValue5 &&
               cellValue5 <= matchValue6;
      case CompareOperator.NOT_BETWEEN:
        var values2 = matchValue.Split(StaticSettings.DefaultNumberStringSplitCharacter);
        if (values2.Length != 2) return false;

        return double.TryParse(checkCellValue, CultureInfo.InvariantCulture, out var cellValue6) &&
               double.TryParse(values2[0], CultureInfo.InvariantCulture, out var matchValue7) &&
               double.TryParse(values2[1], CultureInfo.InvariantCulture, out var matchValue8) &&
               (cellValue6 < matchValue7 ||
                cellValue6 > matchValue8);
      case CompareOperator.IS_NULL_OR_BLANK:
        return string.IsNullOrWhiteSpace(checkCellValue);
      case CompareOperator.IS_NOT_NULL_OR_BLANK:
        return !string.IsNullOrWhiteSpace(checkCellValue);
      default:
        throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null);
    }
  }

  internal static string? GetNewCellValue(string? cellValue, string? setValue, UpdateOperator setOperator) {
    var isRequiredToParse = setOperator is UpdateOperator.MULTIPLY or UpdateOperator.DIVIDE or UpdateOperator.ADD or UpdateOperator.SUBTRACT;
    double? parsedOldValue = null;
    double? parsedNewValue = null;
    if (isRequiredToParse) {
      if (!double.TryParse(cellValue, CultureInfo.InvariantCulture, out var oldValueDouble)) {
        return cellValue;
      }

      if (!double.TryParse(setValue, CultureInfo.InvariantCulture, out var newValueDouble)) {
        return cellValue;
      }

      parsedOldValue = oldValueDouble;
      parsedNewValue = newValueDouble;
    }

    switch (setOperator) {
      case UpdateOperator.SET:
        return setValue;
      case UpdateOperator.MULTIPLY:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue * parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          return val;
        }

        return cellValue;
      case UpdateOperator.DIVIDE:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue / parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          return val;
        }

        return cellValue;
      case UpdateOperator.ADD:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue + parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          return val;
        }

        return cellValue;
      case UpdateOperator.SUBTRACT:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue - parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          return val;
        }

        return cellValue;
      case UpdateOperator.APPEND:
        return cellValue + setValue;
      case UpdateOperator.PREPEND:
        return setValue + cellValue;
      case UpdateOperator.REPLACE:
        //split the setValue into two parts, the first part is the old value and the second part is the new value
        if (cellValue is null) return cellValue;

        if (setValue is null) return cellValue;

        var values = setValue?.Split(StaticSettings.DefaultReplaceSplitString);
        if (values?.Length != 2) return cellValue;
        return cellValue?.Replace(values[0], values[1]);
      default:
        throw new ArgumentOutOfRangeException(nameof(setOperator), setOperator, null);
    }
  }

  /// <summary>
  /// Validate if the file or directory path exists, then return true if it is a directory.
  /// </summary>
  /// <param name="sourceList"></param>
  /// <returns></returns>
  /// <exception cref="Exception"></exception>
  public static List<string> GetExcelFilesList(string[] sourceList) {
    var array = new List<string>();
    foreach (var source in sourceList) {
      var isFile = File.Exists(source);
      var isDirectory = Directory.Exists(source);
      var isNeither = !isFile && !isDirectory;
      if (isNeither) throw new Exception("File or directory does not exist: " + source);

      if (isDirectory) {
        var files = Directory.GetFiles(source, "*.*", SearchOption.AllDirectories)
                             .Where(s => StaticSettings.SupportedExtensions.Contains(Path.GetExtension(s), StringComparer.OrdinalIgnoreCase));
        array.AddRange(files);
      }
      else {
        array.Add(source);
      }
    }

    return array.Distinct().ToList();
  }

  public static Dictionary<int, string> GetHeadersDictionary(ExcelWorksheet worksheet, int headerRowNumber) {
    var headerRow = worksheet.Cells[headerRowNumber, 1, headerRowNumber, worksheet.Dimension.End.Column];
    if (headerRow is null) throw new Exception("Header is not found in the worksheet.");

    return headerRow.Select((cell, index) => new { cell, index })
                    .ToDictionary(x => x.index, x => x.cell.Text);
  }

  public static void UpdateCellValue(ExcelWorksheet worksheet, int row, int column, string? newValue) {
    worksheet.Cells[row, column].Value = newValue;
  }

  public static bool IsAllMatched(ExcelSimpleData excelSimpleData, int row, FilterRecord[] filters) {
    if (filters.Length == 0) throw new ArgumentException("Filters must be provided when merge operator is AND");
    foreach (var filter in filters)
    foreach (var header in excelSimpleData.Headers) {
      var headerValue = excelSimpleData.Worksheet.Cells[row, header.Key + 1]?.Value?.ToString();
      if (header.Value != filter.Column) continue;

      var res = CheckIfAnyMatch(headerValue, filter.Values, filter.CompareOperator);
      if (!res) return false;
    }

    return true;
  }

  public static bool IsAnyMatched(ExcelSimpleData excelSimpleData, int row, FilterRecord[] filters) {
    if (filters.Length == 0) return true;

    foreach (var filter in filters)
    foreach (var header in excelSimpleData.Headers) {
      if (header.Value != filter.Column) continue;

      var headerValue = excelSimpleData.Worksheet.Cells[row, header.Key + 1]?.Value?.ToString();
      var res = CheckIfAnyMatch(headerValue, filter.Values, filter.CompareOperator);
      if (res) return true;
    }

    return false;
  }

  public static void BackupFile(string file) {
    //Create a backup folder in current directory and add timestamp to the file name and copy
    var backupFolder = Path.Combine(Directory.GetCurrentDirectory(), "backup");
    if (!Directory.Exists(backupFolder)) Directory.CreateDirectory(backupFolder);

    var backupFile = Path.Combine(backupFolder, Path.GetFileNameWithoutExtension(file) + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + Path.GetExtension(file));
    File.Copy(file, backupFile);
  }

  public static SupportedFileType GetFileType(string file) {
    if (file.EndsWith("xml")) return SupportedFileType.XML;

    if (file.EndsWith("json")) return SupportedFileType.JSON;

    if (file.EndsWith("yaml")) return SupportedFileType.YAML;

    throw new Exception("Unsupported file type: " + file);
  }
}