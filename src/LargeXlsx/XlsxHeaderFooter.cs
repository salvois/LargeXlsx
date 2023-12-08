namespace LargeXlsx
{
    public class XlsxHeaderFooter
    {
        public string OddHeader { get; }
        public string OddFooter { get; }
        public string EvenHeader { get; }
        public string EvenFooter { get; }
        public string FirstHeader { get; }
        public string FirstFooter { get; }
        public bool AlignWithMargins { get; }
        public bool ScaleWithDoc { get; }

        public XlsxHeaderFooter(
            string oddHeader = null,
            string oddFooter = null,
            string evenHeader = null,
            string evenFooter = null,
            string firstHeader = null,
            string firstFooter = null,
            bool alignWithMargins = true,
            bool scaleWithDoc = true)
        {
            OddHeader = oddHeader;
            OddFooter = oddFooter;
            EvenHeader = evenHeader;
            EvenFooter = evenFooter;
            FirstHeader = firstHeader;
            FirstFooter = firstFooter;
            AlignWithMargins = alignWithMargins;
            ScaleWithDoc = scaleWithDoc;
        }
    }
}