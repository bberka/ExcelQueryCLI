// See https://aka.ms/new-console-template for more information


using Cocona;
using ExcelQueryCLI;
using Serilog;

Log.Logger = new LoggerConfiguration()
             .MinimumLevel.Verbose()
             .WriteTo.Console()
             .WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Hour)
             .CreateLogger();


var testArgs = new[] {
  "update",
  "-f", "D:\\DataSheet_ItemDataTable_Bartar.xlsm",
  "-s", "Item_Table",
  "--filter-query", "'^Index' GREATER_THAN '800042'", 
  "--set-query", "'~ItemName' SET 'WORKS?'"
};
CoconaApp.Run<ExcelQueryCoconaApp>(testArgs);