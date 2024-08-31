// See https://aka.ms/new-console-template for more information


using Cocona;
using ExcelQueryCLI;
using Serilog;

Log.Logger = new LoggerConfiguration()
             .MinimumLevel.Information()
             .WriteTo.Console()
             .WriteTo.File("logs/log.txt", rollingInterval: RollingInterval.Hour)
             .CreateLogger();

var version = typeof(Program).Assembly.GetName().Version;
Log.Information("ExcelQueryCLI v{version}", version);
#if DEBUG
Log.Information("Debug mode");
// args = [
//   "update",
//   "-f", @"D:\DataSheet_ItemDataTable_Bartar.xlsm",
//   "-s", "Item_Table",
//   "--filter-query", "('^Index') EQUALS ('800017')",
//   "--filter-query", "('^Index') EQUALS ('800001')",
//   "--set-query", "('~ItemName') SET ('WORKS?')",
//   "--set-query", "('ItemClassify') SET ('WORKSXXXXXXXXXXXXXXXX')",
// ];
// args = [
//   "update",
//   "-f", @"D:\QuestDialogDataSheet\DataSheet_QuestTable_NPC_Drigan.xlsm",
//   "-s", "Quest_Table",
//   // "--filter-query", "('^QuestGroup') EQUALS ('3133')",
//   "--filter-query", "('^QuestGroup') NOT_EQUALS ('<주석>')",
//   "--set-query", "('AcceptConditions') SET ('getlevel()>99;')",
//   "--set-query", "('AcceptPushItem') SET ('<null>')",
//   "--set-query", "('CompleteCondition') SET ('checkLevelUp(99);')",
//   "--set-query", "('~CompleteConditionText') SET ('Hit level 99;')",
//   "--set-query", "('BaseReward') SET ('exploreexp(100);')",
//   "--set-query", "('SelectReward') SET ('<null>')",
// ];
args = [
  "update",
  "-f", @"D:\DataSheet_QuestTable_NPC_Calpheon_1.xlsm",
  "-s", "Quest_Table",
  // "--filter-query", "('^QuestGroup') EQUALS ('3133')",
  "--filter-query", "('^QuestGroup') NOT_EQUALS ('<주석>')",
  "--set-query", "('AcceptConditions') SET ('getlevel()>99;')",
  "--set-query", "('AcceptPushItem') SET ('<null>')",
  "--set-query", "('CompleteCondition') SET ('checkLevelUp(99);')",
  "--set-query", "('~CompleteConditionText') SET ('Hit level 99;')",
  "--set-query", "('BaseReward') SET ('exploreexp(100);')",
  "--set-query", "('SelectReward') SET ('<null>')",
];


#endif

CoconaApp.Run<ExcelQueryCoconaApp>(args);