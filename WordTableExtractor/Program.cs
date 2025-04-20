using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;  // 确保导入这个命名空间
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;

class Program
{
    static void Main(string[] args)
    {
        string wordPath = @"D:\code\ものづくり実習\study\WordTableExtractor\test.docx";

        // 自动生成 Excel 文件路径（当前目录 + 时间戳）
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        string excelPath = Path.Combine(Directory.GetCurrentDirectory(), $"WordTable_{timestamp}.xlsx");

        List<List<string>> tableData = ExtractWordTableData(wordPath);
        ExportToExcel(tableData, excelPath);

        Console.WriteLine("处理完成！");
    }

    static List<List<string>> ExtractWordTableData(string wordPath)
    {
        var result = new List<List<string>>();

        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(wordPath, false))
        {
            // 使用 DocumentFormat.OpenXml.Wordprocessing 命名空间中的 Table
            var tables = wordDoc.MainDocumentPart.Document.Body.Elements<DocumentFormat.OpenXml.Wordprocessing.Table>();

            foreach (var table in tables)
            {
                foreach (var row in table.Elements<TableRow>())
                {
                    var rowData = new List<string>();
                    foreach (var cell in row.Elements<TableCell>())
                    {
                        string text = cell.InnerText.Trim();
                        rowData.Add(text);
                    }
                    result.Add(rowData);
                }
            }
        }

        return result;
    }

    static void ExportToExcel(List<List<string>> data, string outputPath)
    {
        using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
        {
            WorkbookPart workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            SheetData sheetData = new SheetData();
            worksheetPart.Worksheet = new Worksheet(sheetData);

            Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());
            Sheet sheet = new Sheet()
            {
                Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "提取表格"
            };
            sheets.Append(sheet);

            foreach (var row in data)
            {
                Row newRow = new Row();
                foreach (var cellText in row)
                {
                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(cellText)
                    };
                    newRow.Append(cell);
                }
                sheetData.Append(newRow);
            }

            workbookPart.Workbook.Save();
        }

        Console.WriteLine($"Excel 文件保存至：{outputPath}");
    }
}
