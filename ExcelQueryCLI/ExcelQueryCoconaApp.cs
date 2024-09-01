﻿using Cocona;
using ExcelQueryCLI.Models;
using ExcelQueryCLI.Models.Delete;
using ExcelQueryCLI.Models.Update;
using ExcelQueryCLI.Static;
using ExcelQueryCLI.Xl;
using OfficeOpenXml;
using Serilog;

namespace ExcelQueryCLI;

public sealed class ExcelQueryCoconaApp
{
  [Command("update", Description = "Update rows in Excel file")]
  public void Update(
    [Option("query", ['q'], Description = "Query file path (YAML, JSON, or XML)")]
    string filePath,
    [Option("parallelism", ['p'], Description = "Number of parallel threads")]
    byte parallelThreads = 1
  ) {
    if (parallelThreads < 1) {
      Log.Error("Parallel threads must be greater than or equal to 1.");
      return;
    }

    ExcelUpdateQuery q;
    try {
      var fileType = ExcelTools.GetFileType(filePath);
      q = ExcelUpdateQuery.ParseFile(filePath, fileType);
    }
    catch (Exception ex) {
      Log.Error("Error parsing query file: {Message}", ex.Message);
      return;
    }

    Log.Information("Processing update query");
    try {
      var manager = new EpPlusExcelQueryManager(parallelThreads);
      manager.RunUpdateQuery(q);
    }
    catch (Exception ex) {
      Log.Error("Error updating Excel file: {Message}", ex.Message);
    }
  }

  [Command("delete", Description = "Delete rows in Excel file")]
  public void Delete(
    [Option("query", ['q'], Description = "Yaml query file path")]
    string filePath,
    [Option("parallelism", ['p'], Description = "Number of parallel threads")]
    byte parallelThreads = 1
  ) {
    if (parallelThreads < 1) {
      Log.Error("Parallel threads must be greater than or equal to 1.");
      return;
    }

    ExcelDeleteQuery q;
    try {
      var fileType = ExcelTools.GetFileType(filePath);
      q = ExcelDeleteQuery.ParseFile(filePath, fileType);
    }
    catch (Exception ex) {
      Log.Error("Error parsing query file: {Message}", ex.Message);
      return;
    }


    Log.Information("Processing delete query");
    try {
      var manager = new EpPlusExcelQueryManager(parallelThreads);
      manager.RunDeleteQuery(q);
    }
    catch (Exception ex) {
      Log.Error("Error deleting Excel file: {Message}", ex.Message);
    }
  }
}