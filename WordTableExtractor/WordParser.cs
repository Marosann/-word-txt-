using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Text.RegularExpressions;

namespace WordTableExtractor;
public static class WordParser
{
    private static readonly string[] TargetSections = new[] { "修正内容", "更新内容", "仕様変更内容" };

    public static List<TableItem> ExtractTableItems(string path)
    {
        var items = new List<TableItem>();

        using (var doc = WordprocessingDocument.Open(path, false))
        {
            var elements = doc.MainDocumentPart.Document.Body.Elements<OpenXmlElement>();

            string currentTitle = null;
            List<string> currentVersions = new();
            bool inTargetSection = false;

            foreach (var elem in elements)
            {
                if (elem is Paragraph para)
                {
                    string text = para.InnerText.Trim();

                    // 判断是否进入了目标章节
                    foreach (var title in TargetSections)
                    {
                        if (text.Contains(title))
                        {
                            currentTitle = title;
                            currentVersions.Clear();
                            inTargetSection = true;
                            break;
                        }
                    }

                    if (!inTargetSection) continue;

                    // 匹配单个或多个版本号：例如“版本13-10和版本13-11”
                    var matches = Regex.Matches(text, @"版本?([0-9]{2}-[0-9]{2}(?:-/[A-Z])?)");
                    if (matches.Count > 0)
                    {
                        currentVersions = matches.Select(m => m.Groups[1].Value).Distinct().ToList();
                    }
                }
                else if (elem is Table table && inTargetSection && currentVersions.Count > 0)
                {
                    foreach (var version in currentVersions)
                    {
                        items.AddRange(ExtractFromTable(table, currentTitle, version));
                    }
                }
            }
        }

        return items;
    }

    private static List<TableItem> ExtractFromTable(Table table, string section, string version)
    {
        var results = new List<TableItem>();

        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().ToList();
            for (int i = 0; i < cells.Count - 1; i++)
            {
                string label = cells[i].InnerText.Trim();
                string value = cells[i + 1].InnerText.Trim();

                if (label.Contains("修正ID"))
                {
                    results.Add(new TableItem
                    {
                        Title = section,
                        Version = version,
                        FixId = value
                    });
                }
                else if (label.Contains("修正詳細") || label.Contains("修正详细"))
                {
                    // 查找最后一个 FixId 对应项，填入详细说明
                    var last = results.LastOrDefault(r => r.Version == version && r.Title == section && string.IsNullOrEmpty(r.FixDetail));
                    if (last != null)
                    {
                        last.FixDetail = value;
                    }
                }
            }
        }

        return results;
    }
}
