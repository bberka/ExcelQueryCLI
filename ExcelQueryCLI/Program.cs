// See https://aka.ms/new-console-template for more information


using Cocona;
using ExcelQueryCLI;
using Serilog;

Log.Logger = new LoggerConfiguration()
             .MinimumLevel.Information()
             .WriteTo.Console()
             .WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Hour)
             .CreateLogger();


#if DEBUG
args = [
  "update",
  "-f", "D:\\DataSheet_ItemDataTable_Bartar.xlsm",
  "-s", "Item_Table",
  // "--filter-query", "('^Index' OR 'ItemType') EQUALS '800001'",
  // "--filter-query", "('^Index' OR 'ItemType') EQUALS ('800001')",
  "--filter-query", "('^Index' OR 'ItemType') EQUALS ('800001' OR '800002' OR '800005')",
  // "--filter-query", "('^Index' OR 'ItemType') EQUALS '800001' OR '800002'",
  "--set-query", "'~ItemName' SET 'WORKS?'",
  // "--set-query", "'ItemType' SET 'XXX'"
];

// args = [
//   "delete",
//   "-f", "D:\\DataSheet_ItemDataTable_Bartar.xlsm",
//   "-s", "Item_Table",
//   "--filter-query", "'^Index' EQUALS '800001'",
//   "--filter-query", "'^Index' EQUALS '800002'",
//   // "--filter-query", "'ItemType' EQUALS '0'"
// ];
#endif

CoconaApp.Run<ExcelQueryCoconaApp>(args);