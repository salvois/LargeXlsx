namespace LargeXlsx
{
    public class XlsxColumn
    {
        public int Count { get; }
        public bool Hidden { get; }
        public XlsxStyle Style { get; }
        public double? Width { get; }

        public static XlsxColumn Unformatted(int count = 1)
        {
            return new XlsxColumn(count, false, null, null);
        }

        public static XlsxColumn Formatted(double width, int count = 1, bool hidden = false, XlsxStyle style = null)
        {
            return new XlsxColumn(count, hidden, style, width);
        }

        private XlsxColumn(int count, bool hidden, XlsxStyle style, double? width)
        {
            Count = count;
            Hidden = hidden;
            Style = style;
            Width = width;
        }
    }
}