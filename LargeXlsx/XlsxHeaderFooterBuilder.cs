using System.Text;

namespace LargeXlsx
{
    public class XlsxHeaderFooterBuilder
    {
        public override string ToString() => _sb.ToString();

        public XlsxHeaderFooterBuilder Left() => DoAppend("&L");
        public XlsxHeaderFooterBuilder Center() => DoAppend("&C");
        public XlsxHeaderFooterBuilder Right() => DoAppend("&R");
        public XlsxHeaderFooterBuilder Text(string text) => DoAppend(text.Replace("&", "&&"));
        public XlsxHeaderFooterBuilder CurrentDate() => DoAppend("&D");
        public XlsxHeaderFooterBuilder CurrentTime() => DoAppend("&T");
        public XlsxHeaderFooterBuilder FileName() => DoAppend("&F");
        public XlsxHeaderFooterBuilder FilePath() => DoAppend("&Z");
        public XlsxHeaderFooterBuilder NumberOfPages() => DoAppend("&N");
        public XlsxHeaderFooterBuilder PageNumber(int offset = 0) => offset == 0 ? DoAppend("&P") : DoAppend($"&P{offset:+0;-0}");
        public XlsxHeaderFooterBuilder SheetName() => DoAppend("&A");
        public XlsxHeaderFooterBuilder FontSize(int points) => DoAppend($"&{points:0}");
        public XlsxHeaderFooterBuilder Font(string name, bool bold = false, bool italic = false) => DoAppend($"&\"{name},{GetFontType(bold, italic)}\"");
        public XlsxHeaderFooterBuilder Font(bool bold = false, bool italic = false) => DoAppend($"&\"-,{GetFontType(bold, italic)}\"");
        public XlsxHeaderFooterBuilder Bold() => DoAppend("&B");
        public XlsxHeaderFooterBuilder Italic() => DoAppend("&I");
        public XlsxHeaderFooterBuilder Underline() => DoAppend("&U");
        public XlsxHeaderFooterBuilder DoubleUnderline() => DoAppend("&E");
        public XlsxHeaderFooterBuilder StrikeThrough() => DoAppend("&S");
        public XlsxHeaderFooterBuilder Subscript() => DoAppend("&Y");
        public XlsxHeaderFooterBuilder Superscript() => DoAppend("&X");

        private readonly StringBuilder _sb = new StringBuilder();

        private XlsxHeaderFooterBuilder DoAppend(string text)
        {
            _sb.Append(text);
            return this;
        }

        private static string GetFontType(bool bold, bool italic)
        {
            if (!bold && !italic) return "Regular";
            if (bold && !italic) return "Bold";
            if (!bold) return "Italic";
            return "Bold Italic";
        }
    }
}