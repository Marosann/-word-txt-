using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WordTableExtractor.Helpers
{
    public static class ExcelExporter
    {
        public static void ExportToExcel(Dictionary<string, List<TableItem>> groupedData, string outputPath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());

                uint sheetId = 1;
                foreach (var version in groupedData.Keys)
                {
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    SheetData sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    Sheet sheet = new Sheet()
                    {
                        Id = workbookPart.GetIdOfPart(worksheetPart),
                        SheetId = sheetId++,
                        Name = SanitizeSheetName(version)
                    };
                    sheets.Append(sheet);

                    // 添加表头
                    Row header = new Row();
                    header.Append(CreateTextCell("修正ID"));
                    header.Append(CreateTextCell("修正詳細"));
                    header.Append(CreateTextCell("バージョン"));
                    header.Append(CreateTextCell("章タイトル"));
                    sheetData.Append(header);

                    // 添加数据
                    foreach (var item in groupedData[version])
                    {
                        Row row = new Row();
                        row.Append(CreateTextCell(item.FixId));
                        row.Append(CreateTextCell(item.FixDetail));
                        row.Append(CreateTextCell(item.Version));
                        row.Append(CreateTextCell(item.Title));
                        sheetData.Append(row);
                    }
                }

                workbookPart.Workbook.Save();
            }
        }

        private static Cell CreateTextCell(string text)
        {
            return new Cell
            {
                DataType = CellValues.String,
                CellValue = new CellValue(text ?? "")
            };
        }

        private static string SanitizeSheetName(string name)
        {
            foreach (var c in new[] { '\\', '/', '*', '[', ']', ':', '?' })
            {
                name = name.Replace(c, '_');
            }
            return name.Length > 31 ? name.Substring(0, 31) : name;
        }
    }
}
