using System.Runtime.CompilerServices;

namespace LargeXlsx
{
    public class XlsxHeaderFooterText
    {
        public string CenteredText { get; }
        public string LeftAlignedText { get; }
        public string RightAlignedText { get; }

        public XlsxHeaderFooterText(
            string leftAlignedText = null,
            string centeredText = null, 
            string rightAlignedText = null)
        {
            CenteredText = centeredText;
            LeftAlignedText = leftAlignedText;
            RightAlignedText = rightAlignedText;
        }

        public string WriteText()
        {
            var text = "";
            if (LeftAlignedText != null)
                text += $"&L{LeftAlignedText}";
            if (CenteredText != null)
                text += $"&C{CenteredText}";
            if (RightAlignedText != null)
                text += $"&R{RightAlignedText}";

            return text;
        }
    }
    
    public class XlsxHeaderFooter
    {
        /// <summary>
        /// Inserts the current date.
        /// </summary>
        public const string CurrentDate = "&D";
        /// <summary>
        /// Inserts the current time.
        /// </summary>
        public const string CurrentTime = "&T";
        /// <summary>
        /// Inserts the name of workbook file.
        /// </summary>
        public const string FileName = "&F";
        /// <summary>
        /// Inserts the workbook file path.
        /// </summary>
        public const string FilePath = "&Z";
        /// <summary>
        /// Inserts the total number of pages in a workbook.
        /// </summary>
        public const string NumberOfPages = "&N";
        /// <summary>
        /// Inserts the current page number.
        /// </summary>
        public const string PageNumber = "&P";
        /// <summary>
        /// Inserts the name of a worksheet.
        /// </summary>
        public const string SheetName = "&A";
        /// <summary>
        /// Turns bold on or off for the characters that follow.
        /// </summary>
        public const string Bold = "&B";
        /// <summary>
        /// Turns italic on or off for the characters that follow.
        /// </summary>
        public const string Italic = "&I";
        /// <summary>
        /// Turns underline on or off for the characters that follow.
        /// </summary>
        public const string Underline = "&U";
        /// <summary>
        /// Turns double underline on or off for the characters that follow.
        /// </summary>
        public const string DoubleUnderline = "&E";
        /// <summary>
        /// Turns strikethrough on or off for the characters that follow.
        /// </summary>
        public const string Strikethrough = "&S";
        /// <summary>
        /// Turns subscript on or off for the characters that follow.
        /// </summary>
        public const string Subscript = "&Y";
        /// <summary>
        /// Turns superscript on or off for the characters that follow.
        /// </summary>
        public const string Superscript = "&X";
        
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
        public XlsxHeaderFooterText EvenFooter { get; }
        public XlsxHeaderFooterText EvenHeader { get; }
        public XlsxHeaderFooterText FirstFooter { get; }
        public XlsxHeaderFooterText FirstHeader { get; }
        public XlsxHeaderFooterText OddFooter { get; }
        public XlsxHeaderFooterText OddHeader { get; }
        
        public XlsxHeaderFooter(
            bool alignWithMargins = true, 
            bool differentFirst = false, 
            bool differentOddEven = false, 
            bool scaleWithDoc = false, 
            XlsxHeaderFooterText evenFooter = null, 
            XlsxHeaderFooterText evenHeader = null, 
            XlsxHeaderFooterText firstFooter = null, 
            XlsxHeaderFooterText firstHeader = null, 
            XlsxHeaderFooterText oddFooter = null, 
            XlsxHeaderFooterText oddHeader = null)
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