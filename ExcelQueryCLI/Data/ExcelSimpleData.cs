using OfficeOpenXml;

namespace ExcelQueryCLI.Data;

public sealed record ExcelSimpleData(
  ExcelWorksheet Worksheet,
  Dictionary<int, string> Headers);