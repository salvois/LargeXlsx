using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LargeXlsx
{
    public class XlsxStylesheet2
    {
        public static readonly XlsxFill NoFill = new XlsxFill(0);
        public static readonly XlsxFill Gray125Fill = new XlsxFill(1);
        public static readonly XlsxFont DefaultFont = new XlsxFont(0);
        public static readonly XlsxNumberFormat GeneralNumberFormat = new XlsxNumberFormat(0);
        public static readonly XlsxNumberFormat TwoDecimalExcelNumberFormat = new XlsxNumberFormat(4); // #,##0.00
        public static readonly XlsxBorder NoBorder = new XlsxBorder(0);
        public static readonly XlsxStyle DefaultStyle = new XlsxStyle(0);

        private readonly Stylesheet _stylesheet;
        private readonly Dictionary<StyleTuple, XlsxStyle> _styles;
        private uint _nextFontId;
        private uint _nextBorderId;
        private uint _nextFillId;
        private uint _nextNumberFormatId;
        private uint _nextStyleId;

        internal XlsxStylesheet2()
        {
            _stylesheet = new Stylesheet
            {
                Fonts = new Fonts(),
                Borders = new Borders(),
                Fills = new Fills(),
                NumberingFormats = new NumberingFormats(),
                CellFormats = new CellFormats()
            };

            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.None } });
            _stylesheet.Fills.AppendChild(new Fill { PatternFill = new PatternFill { PatternType = PatternValues.Gray125 } });
            _nextFillId = 2; // ids less than 2 are reserved by Excel for default fills

            _stylesheet.Borders.AppendChild(new Border {
                TopBorder = new TopBorder(),
                RightBorder = new RightBorder(),
                BottomBorder = new BottomBorder(),
                LeftBorder = new LeftBorder(),
                DiagonalBorder = new DiagonalBorder()
            });
            _nextBorderId = 1;

            _nextNumberFormatId = 164;  // ids less than 164 are reserved by Excel for default formats

            _nextFontId = 0;
            CreateFont("Calibri", 11, "000000");

            _styles = new Dictionary<StyleTuple, XlsxStyle>();
            _nextStyleId = 0;
            CreateStyle(DefaultFont, NoFill, GeneralNumberFormat, NoBorder);
        }

        public XlsxFont CreateFont(string fontName, double fontSize, string hexRgbColor)
        {
            _stylesheet.Fonts.AppendChild(new Font
            {
                FontSize = new FontSize { Val = fontSize },
                Color = new Color { Rgb = HexBinaryValue.FromString(hexRgbColor) },
                FontName = new FontName { Val = fontName },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
                FontScheme = new FontScheme { Val = FontSchemeValues.Minor }
            });
            return new XlsxFont(_nextFontId++);
        }

        public XlsxFill CreateSolidFill(string hexRgbColor)
        {
            var hexBinaryValue = HexBinaryValue.FromString(hexRgbColor);
            _stylesheet.Fills.AppendChild(new Fill
            {
                PatternFill = new PatternFill
                {
                    PatternType = PatternValues.Solid,
                    BackgroundColor = new BackgroundColor { Rgb = hexBinaryValue },
                    ForegroundColor = new ForegroundColor { Rgb = hexBinaryValue }
                }
            });
            return new XlsxFill(_nextFillId++);
        }

        public XlsxNumberFormat CreateNumberFormat(string formatCode)
        {
            _stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = _nextNumberFormatId,
                FormatCode = formatCode
            });
            return new XlsxNumberFormat(_nextNumberFormatId++);
        }

        public XlsxBorder CreateBorder(BorderStyleValues top, BorderStyleValues right, BorderStyleValues bottom, BorderStyleValues left, string hexRgbColor)
        {
            var hexBinaryValue = HexBinaryValue.FromString(hexRgbColor);
            var border = new Border
            {
                TopBorder = new TopBorder { Color = new Color { Rgb = hexBinaryValue }, Style = top },
                RightBorder = new RightBorder { Color = new Color { Rgb = hexBinaryValue }, Style = right },
                BottomBorder = new BottomBorder { Color = new Color { Rgb = hexBinaryValue }, Style = bottom },
                LeftBorder = new LeftBorder { Color = new Color { Rgb = hexBinaryValue }, Style = left },
                DiagonalBorder = new DiagonalBorder()
            };
            _stylesheet.Borders.AppendChild(border);
            return new XlsxBorder(_nextBorderId++);
        }

        public XlsxStyle CreateStyle(XlsxFont font, XlsxFill fill, XlsxNumberFormat numberFormat, XlsxBorder border)
        {
            var styleTuple = new StyleTuple(font.Id, fill.Id, numberFormat.Id, border.Id);
            if (_styles.TryGetValue(styleTuple, out var styleId))
                return styleId;

            _stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = numberFormat.Id,
                FontId = font.Id,
                FillId = fill.Id,
                BorderId = border.Id,
                ApplyFont = true,
                ApplyFill = true,
                ApplyNumberFormat = true,
                ApplyBorder = true
            });
            var newStyle = new XlsxStyle(_nextStyleId++);
            _styles[styleTuple] = newStyle;
            return newStyle;
        }

        internal void Save(SpreadsheetDocument document)
        {
            _stylesheet.Fonts.Count = (uint)_stylesheet.Fonts.ChildElements.Count;
            _stylesheet.Borders.Count = (uint)_stylesheet.Borders.ChildElements.Count;
            _stylesheet.Fills.Count = (uint)_stylesheet.Fills.ChildElements.Count;
            _stylesheet.NumberingFormats.Count = (uint)_stylesheet.NumberingFormats.ChildElements.Count;
            _stylesheet.CellFormats.Count = (uint)_stylesheet.CellFormats.ChildElements.Count;

            var workbookStylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            workbookStylesPart.Stylesheet = _stylesheet;
            _stylesheet.Save();
        }

        private struct StyleTuple
        {
            public uint FontId { get; }
            public uint FillId { get; }
            public uint NumberFormatId { get; }
            public uint BorderId { get; }

            public StyleTuple(uint fontId, uint fillId, uint numberFormatId, uint borderId)
            {
                FontId = fontId;
                FillId = fillId;
                NumberFormatId = numberFormatId;
                BorderId = borderId;
            }
        }
    }
}