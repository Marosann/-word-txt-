using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using WordTableExtractor.Models;
using System.Text.RegularExpressions;

namespace WordTableExtractor.Helpers;

public class WordParser
{
    private static readonly Regex VersionPattern = new(@"\b(\d{2}-\d{2}(?:-/[A-Z])?(?:\s*UPDATE\d+[a-zA-Z]?)?)\b", RegexOptions.Compiled);
    private static readonly string[] ValidCategories = ["機能追加内容", "仕様変更内容", "修正内容", "改善内容"];

    public static List<SimpleTableItem> ParseTargetSections(string wordPath)
    {
        var result = new List<SimpleTableItem>();

        using var wordDoc = WordprocessingDocument.Open(wordPath, false);
        var bodyElements = wordDoc.MainDocumentPart?.Document.Body?.Elements().ToList();
        if (bodyElements == null) return result;

        bool isInTargetSection = false;
        List<string> currentVersions = new();
        string? currentCategory = null;

        foreach (var element in bodyElements)
        {
            if (element is Paragraph para)
            {
                string paraText = para.InnerText.Trim();

                var paraProps = para.ParagraphProperties;
                var styleId = paraProps?.ParagraphStyleId?.Val?.Value;
                if (styleId == "Heading1" && (paraText.StartsWith("6.") || paraText.StartsWith("7.") || paraText.StartsWith("8.")))
                {
                    isInTargetSection = true;
                    continue;
                }
                else if (styleId == "Heading1" && paraText.StartsWith("9."))
                {
                    isInTargetSection = false;
                    continue;
                }

                if (isInTargetSection)
                {
                    // 提取多个版本号
                    var matches = VersionPattern.Matches(paraText);
                    if (matches.Count > 0)
                    {
                        currentVersions = matches.Select(m => m.Value).ToList();
                    }
                }
            }
            else if (element is Table table && isInTargetSection)
            {
                var allCells = table.Descendants<TableCell>().Select(c => c.InnerText.Trim()).ToList();

                foreach (var category in ValidCategories)
                {
                    if (allCells.Any(cell => cell.Contains(category)))
                    {
                        currentCategory = category;
                        break;
                    }
                }

                if (currentCategory != null && currentVersions.Any())
                {
                    foreach (var version in currentVersions)
                    {
                        var item = new SimpleTableItem
                        {
                            Versions = [version],
                            Category = currentCategory,
                            FixId = TryExtractFollowingValue(table, "修正ID"),
                            FixDetail = TryExtractFollowingValue(table, "修正詳細")
                        };
                        result.Add(item);
                    }
                }
            }
        }

        return result;
    }

    private static string? TryExtractFollowingValue(Table table, string label)
    {
        foreach (var row in table.Elements<TableRow>())
        {
            var cells = row.Elements<TableCell>().Select(c => c.InnerText.Trim()).ToList();
            for (int i = 0; i < cells.Count - 1; i++)
            {
                if (cells[i].Contains(label))
                {
                    return cells[i + 1];
                }
            }
        }
        return null;
    }
}
