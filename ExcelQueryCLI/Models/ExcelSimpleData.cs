using OfficeOpenXml;

namespace ExcelQueryCLI.Models;

public sealed record ExcelSimpleData(
  ExcelWorksheet Worksheet,
  Dictionary<int, string> Headers);