using System.Globalization;
using ExcelQueryCLI.Parsers;
using ExcelQueryCLI.Static;
using Serilog;

namespace ExcelQueryCLI;

internal static class ExcelTools
{
  internal static bool CheckIfMatchingFilter(string? checkCellValue, List<string> matchValues, FilterOperator @operator) {
    //Check if matches any
    foreach (var checkValue in matchValues) {
      if (CheckIfMatchingFilter(checkCellValue, checkValue, @operator)) {
        return true;
      }
    }

    return false;
  }

  internal static bool CheckIfMatchingFilter(string? checkCellValue, string matchValue, FilterOperator @operator) {
    Log.Verbose("CheckFilter::Cell Value: {cellFilterValue}, Match Value: {matchFilterValue}, Operator: {Operator}",
                checkCellValue,
                matchValue,
                @operator);

    if (checkCellValue is null) {
      return false;
    }

    switch (@operator) {
      case FilterOperator.EQUALS:
        return checkCellValue == matchValue;
      case FilterOperator.NOT_EQUALS:
        return checkCellValue != matchValue;
      case FilterOperator.GREATER_THAN:
        return double.TryParse(checkCellValue, out var cellValue1) && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue1) && cellValue1 > matchValue1;
      case FilterOperator.LESS_THAN:
        return double.TryParse(checkCellValue, out var cellValue2) && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue2) && cellValue2 < matchValue2;
      case FilterOperator.GREATER_THAN_OR_EQUAL:
        return double.TryParse(checkCellValue, out var cellValue3) && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue3) && cellValue3 >= matchValue3;
      case FilterOperator.LESS_THAN_OR_EQUAL:
        return double.TryParse(checkCellValue, out var cellValue4) && double.TryParse(matchValue, CultureInfo.InvariantCulture, out var matchValue4) && cellValue4 <= matchValue4;
      case FilterOperator.CONTAINS:
        return checkCellValue.Contains(matchValue);
      case FilterOperator.NOT_CONTAINS:
        return !checkCellValue.Contains(matchValue);
      case FilterOperator.STARTS_WITH:
        return checkCellValue.StartsWith(matchValue);
      case FilterOperator.ENDS_WITH:
        return checkCellValue.EndsWith(matchValue);
      case FilterOperator.BETWEEN:
        var values = matchValue.Split("<>");
        if (values.Length != 2) {
          return false;
        }

        return double.TryParse(checkCellValue, CultureInfo.InvariantCulture, out var cellValue5) &&
               double.TryParse(values[0], CultureInfo.InvariantCulture, out var matchValue5) &&
               double.TryParse(values[1], CultureInfo.InvariantCulture, out var matchValue6) &&
               cellValue5 >= matchValue5 &&
               cellValue5 <= matchValue6;
      case FilterOperator.NOT_BETWEEN:
        var values2 = matchValue.Split("|");
        if (values2.Length != 2) {
          return false;
        }

        return double.TryParse(checkCellValue, CultureInfo.InvariantCulture, out var cellValue6) &&
               double.TryParse(values2[0], CultureInfo.InvariantCulture, out var matchValue7) &&
               double.TryParse(values2[1], CultureInfo.InvariantCulture, out var matchValue8) &&
               (cellValue6 < matchValue7 ||
                cellValue6 > matchValue8);
      default:
        throw new ArgumentOutOfRangeException(nameof(@operator), @operator, null);
    }
  }

  internal static string? GetNewCellValue(string? cellValue, string? setValue, UpdateOperator setOperator) {
    Log.Verbose("UpdateCellValue::Cell Value: {cellValue}, Set Value: {setValue}, Operator: {setOperator}",
                cellValue,
                setValue,
                setOperator);

    var isRequiredToParse = setOperator == UpdateOperator.MULTIPLY ||
                            setOperator == UpdateOperator.DIVIDE ||
                            setOperator == UpdateOperator.ADD ||
                            setOperator == UpdateOperator.SUBTRACT;


    double? parsedOldValue = null;
    double? parsedNewValue = null;
    if (isRequiredToParse) {
      if (!double.TryParse(cellValue, CultureInfo.InvariantCulture, out var oldValueDouble)) {
        Log.Verbose("UpdateCellValue::Failed to parse old value: {oldValue}", cellValue);
        return cellValue;
      }

      if (!double.TryParse(setValue, CultureInfo.InvariantCulture, out var newValueDouble)) {
        Log.Verbose("UpdateCellValue::Failed to parse new value: {newValue}", setValue);
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
          Log.Verbose("UpdateCellValue::Multiplying {oldValue} by {newValue} = {result}", parsedOldValue, parsedNewValue, val);
          return val;
        }

        Log.Verbose("UpdateCellValue::Failed to parse old or new value: {oldValue} {newValue}", cellValue, setValue);
        return cellValue;
      case UpdateOperator.DIVIDE:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue / parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          Log.Verbose("UpdateCellValue::Dividing {oldValue} by {newValue} = {result}", parsedOldValue, parsedNewValue, val);
          return val;
        }

        Log.Verbose("UpdateCellValue::Failed to parse old or new value: {oldValue} {newValue}", cellValue, setValue);
        return cellValue;
      case UpdateOperator.ADD:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue + parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          Log.Verbose("UpdateCellValue::Adding {oldValue} by {newValue} = {result}", parsedOldValue, parsedNewValue, val);
          return val;
        }

        Log.Verbose("UpdateCellValue::Failed to parse old or new value: {oldValue} {newValue}", cellValue, setValue);
        return cellValue;
      case UpdateOperator.SUBTRACT:
        if (parsedOldValue.HasValue && parsedNewValue.HasValue) {
          var val = ((double)(parsedOldValue - parsedNewValue)).ToString(CultureInfo.InvariantCulture);
          Log.Verbose("UpdateCellValue::Subtracting {oldValue} by {newValue} = {result}", parsedOldValue, parsedNewValue, val);
          return val;
        }

        Log.Verbose("UpdateCellValue::Failed to parse old or new value: {oldValue} {newValue}", cellValue, setValue);
        return cellValue;
      case UpdateOperator.APPEND:
        return cellValue + setValue;
      case UpdateOperator.PREPEND:
        return setValue + cellValue;
      case UpdateOperator.REPLACE:
        //split the setValue into two parts, the first part is the old value and the second part is the new value
        var values = setValue?.Split("<>");
        if (values?.Length != 2) {
          return cellValue;
        }

        return cellValue?.Replace(values[0], values[1]) ?? cellValue;
      default:
        throw new ArgumentOutOfRangeException(nameof(setOperator), setOperator, null);
    }
  }
  internal static Dictionary<int, SetQueryParser> GetSetQueryColumnIndexDict(List<SetQueryParser> setQueries, List<string> headers) {
    var result = new Dictionary<int, SetQueryParser>();
    foreach (var setQuery in setQueries) {
      var setColumnIndex = headers.FindIndex(header => header.Equals(setQuery.Column, StringComparison.OrdinalIgnoreCase));
      if (setColumnIndex == -1) {
        Log.Warning("Set column {setQuery.Column} not found.", setQuery.Column);
        continue;
      }

      result[setColumnIndex] = setQuery;
    }

    return result;
  }

  internal static List<Tuple<List<int>, FilterQueryParser>>? GetFilterQueryColumnIndexTuple(List<FilterQueryParser>? filterQueries, List<string> headers) {
    List<Tuple<List<int>, FilterQueryParser>>? result = null;
    if (filterQueries is null) return result;
    result = new List<Tuple<List<int>, FilterQueryParser>>();
    foreach (var filterQuery in filterQueries) {
      var indexes = new List<int>();
      foreach (var col in filterQuery.Columns) {
        var filterColumnIndex = headers.FindIndex(header => header.Equals(col, StringComparison.OrdinalIgnoreCase));
        if (filterColumnIndex == -1) {
          Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
          continue;
        }

        indexes.Add(filterColumnIndex);
      }

      if (indexes.Count == 0) {
        Log.Warning("Filter column {filterQuery.Column} not found.", filterQuery.Columns);
        continue;
      }

      result.Add(new Tuple<List<int>, FilterQueryParser>(indexes, filterQuery));
    }

    return result;
  }
}