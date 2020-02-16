/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020 Salvatore ISAJA. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED THE COPYRIGHT HOLDER ``AS IS'' AND ANY EXPRESS
OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN
NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LargeXlsx
{
    /*
     * Special thanks to http://polymathprogrammer.com/2009/11/09/how-to-create-stylesheet-in-excel-open-xml/
     * for very valuable insights on how to properly create styles.
     */
    public class XlsxStylesheet
    {
        private readonly Stylesheet _stylesheet;
        private readonly Dictionary<StyleTuple, XlsxStyle> _styles;
        private uint _nextFontId;
        private uint _nextBorderId;
        private uint _nextFillId;
        private uint _nextNumberFormatId;
        private uint _nextStyleId;

        internal XlsxStylesheet()
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
            CreateStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.General);
        }

        public XlsxFont CreateFont(string fontName, double fontSize, string hexRgbColor, bool bold = false, bool italic = false, bool strike = false)
        {
            var font = new Font
            {
                FontSize = new FontSize { Val = fontSize },
                Color = new Color { Rgb = HexBinaryValue.FromString(hexRgbColor) },
                FontName = new FontName { Val = fontName },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
                FontScheme = new FontScheme { Val = FontSchemeValues.Minor }
            };
            if (bold) font.Bold = new Bold();
            if (italic) font.Italic = new Italic();
            if (strike) font.Strike = new Strike();
            _stylesheet.Fonts.AppendChild(font);
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

        public XlsxBorder CreateBorder(
            string hexRgbColor,
            BorderStyleValues top = BorderStyleValues.None,
            BorderStyleValues right = BorderStyleValues.None,
            BorderStyleValues bottom = BorderStyleValues.None,
            BorderStyleValues left = BorderStyleValues.None)
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

        public XlsxStyle CreateStyle(XlsxFont font, XlsxFill fill, XlsxBorder border, XlsxNumberFormat numberFormat)
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

    public struct XlsxFont
    {
        public static readonly XlsxFont Default = new XlsxFont(0);

        internal uint Id { get; }

        internal XlsxFont(uint id)
        {
            Id = id;
        }
    }

    public struct XlsxFill
    {
        public static readonly XlsxFill None = new XlsxFill(0);
        public static readonly XlsxFill Gray125 = new XlsxFill(1);

        internal uint Id { get; }

        internal XlsxFill(uint id)
        {
            Id = id;
        }
    }

    public struct XlsxBorder
    {
        public static readonly XlsxBorder None = new XlsxBorder(0);

        internal uint Id { get; }

        internal XlsxBorder(uint id)
        {
            Id = id;
        }
    }

    public struct XlsxNumberFormat
    {
        public static readonly XlsxNumberFormat General = new XlsxNumberFormat(0);
        public static readonly XlsxNumberFormat TwoDecimal = new XlsxNumberFormat(4); // #,##0.00

        internal uint Id { get; }

        internal XlsxNumberFormat(uint id)
        {
            Id = id;
        }
    }

    public struct XlsxStyle
    {
        public static readonly XlsxStyle Default = new XlsxStyle(0);

        internal uint Id { get; }

        internal XlsxStyle(uint id)
        {
            Id = id;
        }
    }
}