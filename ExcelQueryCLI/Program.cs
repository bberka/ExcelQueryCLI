using Cocona;
using ExcelQueryCLI;
using OfficeOpenXml;
using Serilog;
using Serilog.Formatting.Compact;

Log.Logger = new LoggerConfiguration()
             .MinimumLevel.Information()
             .WriteTo.Console()
             .WriteTo.File(new CompactJsonFormatter(), "logs/log.json", rollingInterval: RollingInterval.Hour)
             .CreateLogger();
var version = typeof(Program).Assembly.GetName().Version;
Log.Information("ExcelQueryCLI v{version}", version);

#if DEBUG
args = ["update", "-q", @"D:\test.xml"];
#endif
ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
CoconaApp.Run<ExcelQueryCoconaApp>(args);