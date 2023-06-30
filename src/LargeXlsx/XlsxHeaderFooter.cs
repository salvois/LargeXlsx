namespace LargeXlsx
{
    public class XlsxHeaderFooter
    {
        /// <summary>
        /// Align header footer margins with page margins.
        /// </summary>
        public bool AlignWithMargins { get; }
        /// <summary>
        /// Different first page header and footer.
        /// </summary>
        public bool DifferentFirst { get; }
        /// <summary>
        /// Different odd and even page headers and footers.
        /// </summary>
        public bool DifferentOddEven { get; }
        /// <summary>
        /// Scale header and footer with document scaling.
        /// </summary>
        public bool ScaleWithDoc { get; }
        public string EvenFooter { get; }
        public string EvenHeader { get; }
        public string FirstFooter { get; }
        public string FirstHeader { get; }
        public string OddFooter { get; }
        public string OddHeader { get; }
        
        public XlsxHeaderFooter(
            bool alignWithMargins = true, 
            bool differentFirst = false, 
            bool differentOddEven = false, 
            bool scaleWithDoc = false, 
            string evenFooter = null, 
            string evenHeader = null, 
            string firstFooter = null, 
            string firstHeader = null, 
            string oddFooter = null, 
            string oddHeader = null)
        {
            AlignWithMargins = alignWithMargins;
            DifferentFirst = differentFirst;
            DifferentOddEven = differentOddEven;
            ScaleWithDoc = scaleWithDoc;
            EvenFooter = evenFooter;
            EvenHeader = evenHeader;
            FirstFooter = firstFooter;
            FirstHeader = firstHeader;
            OddFooter = oddFooter;
            OddHeader = oddHeader;
        }
    }
}