using Cocona;
using ExcelQueryCLI;

// Log.Logger = new LoggerConfiguration()
//              .MinimumLevel.Information()
//              .WriteTo.Console()
//              .WriteTo.File(new CompactJsonFormatter(), "logs/log.json", rollingInterval: RollingInterval.Hour)
//              .CreateLogger();
// var version = typeof(Program).Assembly.GetName().Version;
// Log.Information("ExcelQueryCLI v{version}", version);

// #if DEBUG
// args = ["update", "-q", @"D:\test.xml"];
// #endif


var app = CoconaApp.Create();

CoconaApp.Run<ExcelQueryCoconaApp>(args);