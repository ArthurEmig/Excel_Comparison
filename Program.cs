using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace ExcelCompare
{
    class Program
    {
        static int Main(string[] args)
        {
            try
            {
                var options = ParseArgs(args);
                if (options == null)
                {
                    PrintUsage();
                    return 1;
                }

                Console.WriteLine($"File 1: {options.File1}");
                Console.WriteLine($"File 2: {options.File2}");
                Console.WriteLine($"Sheet 1: {options.Sheet1 ?? "(first sheet)"}");
                Console.WriteLine($"Sheet 2: {options.Sheet2 ?? "(first sheet)"}");
                Console.WriteLine($"Output:  {options.Output}");
                Console.WriteLine();

                if (!File.Exists(options.File1))
                {
                    Console.Error.WriteLine($"ERROR: File not found: {options.File1}");
                    return 2;
                }
                if (!File.Exists(options.File2))
                {
                    Console.Error.WriteLine($"ERROR: File not found: {options.File2}");
                    return 2;
                }

                Compare(options);
                Console.WriteLine($"Done. Report written to: {options.Output}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"ERROR: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
                return 99;
            }
        }

        static void PrintUsage()
        {
            Console.WriteLine("Excel Sheet Comparison Tool");
            Console.WriteLine("============================");
            Console.WriteLine();
            Console.WriteLine("Usage:");
            Console.WriteLine("  ExcelCompare.exe <file1.xlsx> <file2.xlsx> [options]");
            Console.WriteLine();
            Console.WriteLine("Options:");
            Console.WriteLine("  --sheet1 <name>    Name of the sheet in file1 (default: first sheet)");
            Console.WriteLine("  --sheet2 <name>    Name of the sheet in file2 (default: first sheet)");
            Console.WriteLine("  --sheet   <name>   Shortcut to set both --sheet1 and --sheet2");
            Console.WriteLine("  --out     <path>   Output .xlsx path (default: Comparison_<timestamp>.xlsx)");
            Console.WriteLine();
            Console.WriteLine("Example:");
            Console.WriteLine("  ExcelCompare.exe old.xlsx new.xlsx --sheet \"Data\" --out report.xlsx");
            Console.WriteLine();
            Console.WriteLine("Output legend:");
            Console.WriteLine("  Yellow = cell value changed");
            Console.WriteLine("  Green  = row/cell added in file2");
            Console.WriteLine("  Red    = row/cell removed (only in file1)");
        }

        class Options
        {
            public string File1 = "";
            public string File2 = "";
            public string? Sheet1;
            public string? Sheet2;
            public string Output = "";
        }

        static Options? ParseArgs(string[] args)
        {
            if (args.Length < 2) return null;
            var opt = new Options
            {
                File1 = args[0],
                File2 = args[1]
            };

            for (int i = 2; i < args.Length; i++)
            {
                string a = args[i].ToLowerInvariant();
                string? next = (i + 1 < args.Length) ? args[i + 1] : null;
                switch (a)
                {
                    case "--sheet1":
                        opt.Sheet1 = next; i++; break;
                    case "--sheet2":
                        opt.Sheet2 = next; i++; break;
                    case "--sheet":
                        opt.Sheet1 = next; opt.Sheet2 = next; i++; break;
                    case "--out":
                    case "--output":
                        opt.Output = next ?? ""; i++; break;
                    default:
                        Console.Error.WriteLine($"Unknown option: {args[i]}");
                        return null;
                }
            }

            if (string.IsNullOrWhiteSpace(opt.Output))
            {
                opt.Output = $"Comparison_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
            }

            return opt;
        }

        static void Compare(Options opt)
        {
            using var wb1 = new XLWorkbook(opt.File1);
            using var wb2 = new XLWorkbook(opt.File2);

            var ws1 = GetSheet(wb1, opt.Sheet1, opt.File1);
            var ws2 = GetSheet(wb2, opt.Sheet2, opt.File2);

            // Determine the bounding box we need to examine
            var used1 = ws1.RangeUsed();
            var used2 = ws2.RangeUsed();

            int lastRow = Math.Max(
                used1?.RangeAddress.LastAddress.RowNumber ?? 0,
                used2?.RangeAddress.LastAddress.RowNumber ?? 0);
            int lastCol = Math.Max(
                used1?.RangeAddress.LastAddress.ColumnNumber ?? 0,
                used2?.RangeAddress.LastAddress.ColumnNumber ?? 0);

            if (lastRow == 0 || lastCol == 0)
            {
                Console.WriteLine("Both sheets are empty — nothing to compare.");
                WriteEmptyReport(opt.Output);
                return;
            }

            using var report = new XLWorkbook();
            var outWs = report.Worksheets.Add("Comparison");

            // Header
            outWs.Cell(1, 1).Value = $"Comparison: {Path.GetFileName(opt.File1)} [{ws1.Name}]  vs  {Path.GetFileName(opt.File2)} [{ws2.Name}]";
            outWs.Cell(1, 1).Style.Font.Bold = true;
            outWs.Cell(1, 1).Style.Font.FontSize = 14;
            outWs.Range(1, 1, 1, Math.Max(4, lastCol)).Merge();

            // Legend
            outWs.Cell(2, 1).Value = "Legend:";
            outWs.Cell(2, 1).Style.Font.Bold = true;
            outWs.Cell(2, 2).Value = "Changed";
            outWs.Cell(2, 2).Style.Fill.BackgroundColor = XLColor.Yellow;
            outWs.Cell(2, 3).Value = "Added (only in file2)";
            outWs.Cell(2, 3).Style.Fill.BackgroundColor = XLColor.LightGreen;
            outWs.Cell(2, 4).Value = "Removed (only in file1)";
            outWs.Cell(2, 4).Style.Fill.BackgroundColor = XLColor.LightPink;

            const int headerOffset = 4; // data rows start at row 4 of the report
            int changes = 0, added = 0, removed = 0;

            // Copy/compare every cell in the bounding box
            for (int r = 1; r <= lastRow; r++)
            {
                for (int c = 1; c <= lastCol; c++)
                {
                    var cell1 = ws1.Cell(r, c);
                    var cell2 = ws2.Cell(r, c);
                    string v1 = CellString(cell1);
                    string v2 = CellString(cell2);

                    var outCell = outWs.Cell(r + headerOffset, c);

                    if (v1 == v2)
                    {
                        // Identical — copy the value (from file2) without highlight
                        if (!string.IsNullOrEmpty(v2))
                            outCell.Value = v2;
                    }
                    else if (string.IsNullOrEmpty(v1) && !string.IsNullOrEmpty(v2))
                    {
                        // Added in file2
                        outCell.Value = v2;
                        outCell.Style.Fill.BackgroundColor = XLColor.LightGreen;
                        outCell.CreateComment().AddText($"Added in file2 (was empty in file1)");
                        added++;
                    }
                    else if (!string.IsNullOrEmpty(v1) && string.IsNullOrEmpty(v2))
                    {
                        // Removed from file2
                        outCell.Value = v1;
                        outCell.Style.Fill.BackgroundColor = XLColor.LightPink;
                        outCell.Style.Font.Strikethrough = true;
                        outCell.CreateComment().AddText($"Removed: was '{v1}' in file1, empty in file2");
                        removed++;
                    }
                    else
                    {
                        // Changed
                        outCell.Value = v2;
                        outCell.Style.Fill.BackgroundColor = XLColor.Yellow;
                        outCell.CreateComment().AddText($"Changed: '{v1}' -> '{v2}'");
                        changes++;
                    }
                }
            }

            // Summary sheet
            var sumWs = report.Worksheets.Add("Summary");
            sumWs.Cell(1, 1).Value = "Metric"; sumWs.Cell(1, 2).Value = "Count";
            sumWs.Cell(1, 1).Style.Font.Bold = true;
            sumWs.Cell(1, 2).Style.Font.Bold = true;
            sumWs.Cell(2, 1).Value = "Cells changed"; sumWs.Cell(2, 2).Value = changes;
            sumWs.Cell(3, 1).Value = "Cells added";   sumWs.Cell(3, 2).Value = added;
            sumWs.Cell(4, 1).Value = "Cells removed"; sumWs.Cell(4, 2).Value = removed;
            sumWs.Cell(5, 1).Value = "Total differences"; sumWs.Cell(5, 2).Value = changes + added + removed;
            sumWs.Cell(5, 1).Style.Font.Bold = true;
            sumWs.Cell(5, 2).Style.Font.Bold = true;
            sumWs.Columns().AdjustToContents();

            outWs.Columns().AdjustToContents();
            outWs.SheetView.FreezeRows(headerOffset);

            report.SaveAs(opt.Output);

            Console.WriteLine($"Changed: {changes}");
            Console.WriteLine($"Added:   {added}");
            Console.WriteLine($"Removed: {removed}");
            Console.WriteLine($"Total:   {changes + added + removed}");
        }

        static IXLWorksheet GetSheet(XLWorkbook wb, string? name, string fileLabel)
        {
            if (string.IsNullOrWhiteSpace(name))
                return wb.Worksheets.First();

            var ws = wb.Worksheets.FirstOrDefault(w =>
                string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase));
            if (ws == null)
            {
                var available = string.Join(", ", wb.Worksheets.Select(w => $"\"{w.Name}\""));
                throw new InvalidOperationException(
                    $"Sheet '{name}' not found in {fileLabel}. Available sheets: {available}");
            }
            return ws;
        }

        static string CellString(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty()) return "";
            try { return cell.GetFormattedString() ?? ""; }
            catch { return cell.Value.ToString() ?? ""; }
        }

        static void WriteEmptyReport(string path)
        {
            using var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Comparison");
            ws.Cell(1, 1).Value = "Both sheets are empty.";
            wb.SaveAs(path);
        }
    }
}
