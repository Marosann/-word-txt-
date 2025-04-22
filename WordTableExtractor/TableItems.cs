namespace WordTableExtractor
{
    public class TableItem
    {
        public string Title { get; set; }            // 主标题，如“修正内容”
        public string Version { get; set; }          // 版本号，如“13-10-/B”
        public string FixId { get; set; }            // 修正ID
        public string FixDetail { get; set; }        // 修正详细
    }
}
