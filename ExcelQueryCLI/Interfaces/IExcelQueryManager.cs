using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Update;

namespace ExcelQueryCLI.Interfaces;

public interface IExcelQueryManager
{
  public void RunUpdateQuery(ExcelUpdateQuery query);
}