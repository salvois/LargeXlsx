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
    
    public class XlsxHeaderFooterSettings
    {
        public static readonly XlsxHeaderFooterSettings Default = new XlsxHeaderFooterSettings(true, false);

        public XlsxHeaderFooterSettings(bool alignWithMargins, bool scaleWithDoc)
        {
            AlignWithMargins = alignWithMargins;
            ScaleWithDoc = scaleWithDoc;
        }

        /// <summary>
        /// Align header footer margins with page margins.
        /// </summary>
        public bool AlignWithMargins { get; }
        /// <summary>
        /// Different first page header and footer.
        /// </summary>
        public bool DifferentFirst { get; set; }
        /// <summary>
        /// Different odd and even page headers and footers.
        /// </summary>
        public bool DifferentOddEven { get; set; }
        /// <summary>
        /// Scale header and footer with document scaling.
        /// </summary>
        public bool ScaleWithDoc { get; }
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
        
        public XlsxHeaderFooterText EvenFooter { get; }
        public XlsxHeaderFooterText EvenHeader { get; }
        public XlsxHeaderFooterText FirstFooter { get; }
        public XlsxHeaderFooterText FirstHeader { get; }
        public XlsxHeaderFooterText OddFooter { get; }
        public XlsxHeaderFooterText OddHeader { get; }
        public XlsxHeaderFooterSettings Settings { get; }
        
        public XlsxHeaderFooter(
            XlsxHeaderFooterText header = null,
            XlsxHeaderFooterText footer = null,
            XlsxHeaderFooterText firstHeader = null,
            XlsxHeaderFooterText firstFooter = null,
            XlsxHeaderFooterText evenHeader = null,
            XlsxHeaderFooterText evenFooter = null, 
            XlsxHeaderFooterSettings settings = null)
        {
            EvenFooter = evenFooter;
            EvenHeader = evenHeader;
            FirstFooter = firstFooter;
            FirstHeader = firstHeader;
            OddFooter = footer;
            OddHeader = header;
            
            if (settings == null)
                settings = XlsxHeaderFooterSettings.Default;

            Settings = settings;
            Settings.DifferentFirst = FirstHeader != null || FirstFooter != null;
            Settings.DifferentOddEven = EvenHeader != null || EvenFooter != null;
        }

        public XlsxHeaderFooter WithHeader(XlsxHeaderFooterText header) =>
            new XlsxHeaderFooter(header, OddFooter, FirstHeader, FirstFooter, EvenHeader, EvenFooter, Settings);
        public XlsxHeaderFooter WithFooter(XlsxHeaderFooterText footer) =>
            new XlsxHeaderFooter(OddHeader, footer, FirstHeader, FirstFooter, EvenHeader, EvenFooter, Settings);
        public XlsxHeaderFooter WithFirstHeader(XlsxHeaderFooterText firstHeader) =>
            new XlsxHeaderFooter(OddHeader, OddFooter, firstHeader, FirstFooter, EvenHeader, EvenFooter, Settings);
        public XlsxHeaderFooter WithFirstFooter(XlsxHeaderFooterText firstFooter) =>
            new XlsxHeaderFooter(OddHeader, OddFooter, FirstHeader, firstFooter, EvenHeader, EvenFooter, Settings);
        public XlsxHeaderFooter WithEvenHeader(XlsxHeaderFooterText evenHeader) =>
            new XlsxHeaderFooter(OddHeader, OddFooter, FirstHeader, FirstFooter, evenHeader, EvenFooter, Settings);
        public XlsxHeaderFooter WithEvenFooter(XlsxHeaderFooterText evenFooter) =>
            new XlsxHeaderFooter(OddHeader, OddFooter, FirstHeader, FirstFooter, EvenHeader, evenFooter, Settings);
    }
}