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
    public class LargeXlsxStylesheet
    {
        public static readonly LargeXlsxFill NoFill = new LargeXlsxFill(0);
        public static readonly LargeXlsxFill Gray125Fill = new LargeXlsxFill(1);
        public static readonly LargeXlsxFont DefaultFont = new LargeXlsxFont(0);
        public static readonly LargeXlsxNumberFormat GeneralNumberFormat = new LargeXlsxNumberFormat(0);
        public static readonly LargeXlsxNumberFormat TwoDecimalExcelNumberFormat = new LargeXlsxNumberFormat(4); // #,##0.00
        public static readonly LargeXlsxBorder NoBorder = new LargeXlsxBorder(0);
        public static readonly LargeXlsxStyle DefaultStyle = new LargeXlsxStyle(0);

        private readonly Stylesheet _stylesheet;
        private readonly Dictionary<StyleTuple, LargeXlsxStyle> _styles;
        private uint _nextFontId;
        private uint _nextBorderId;
        private uint _nextFillId;
        private uint _nextNumberFormatId;
        private uint _nextStyleId;

        internal LargeXlsxStylesheet()
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

            _styles = new Dictionary<StyleTuple, LargeXlsxStyle>();
            _nextStyleId = 0;
            CreateStyle(DefaultFont, NoFill, GeneralNumberFormat, NoBorder);
        }

        public LargeXlsxFont CreateFont(string fontName, double fontSize, string hexRgbColor, bool bold = false, bool italic = false, bool strike = false)
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
            return new LargeXlsxFont(_nextFontId++);
        }

        public LargeXlsxFill CreateSolidFill(string hexRgbColor)
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
            return new LargeXlsxFill(_nextFillId++);
        }

        public LargeXlsxNumberFormat CreateNumberFormat(string formatCode)
        {
            _stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = _nextNumberFormatId,
                FormatCode = formatCode
            });
            return new LargeXlsxNumberFormat(_nextNumberFormatId++);
        }

        public LargeXlsxBorder CreateBorder(BorderStyleValues top, BorderStyleValues right, BorderStyleValues bottom, BorderStyleValues left, string hexRgbColor)
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
            return new LargeXlsxBorder(_nextBorderId++);
        }

        public LargeXlsxStyle CreateStyle(LargeXlsxFont font, LargeXlsxFill fill, LargeXlsxNumberFormat numberFormat, LargeXlsxBorder border)
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
            var newStyle = new LargeXlsxStyle(_nextStyleId++);
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

    public struct LargeXlsxFont
    {
        internal uint Id { get; }

        internal LargeXlsxFont(uint id)
        {
            Id = id;
        }
    }

    public struct LargeXlsxFill
    {
        internal uint Id { get; }

        internal LargeXlsxFill(uint id)
        {
            Id = id;
        }
    }

    public struct LargeXlsxBorder
    {
        internal uint Id { get; }

        internal LargeXlsxBorder(uint id)
        {
            Id = id;
        }
    }

    public struct LargeXlsxNumberFormat
    {
        internal uint Id { get; }

        internal LargeXlsxNumberFormat(uint id)
        {
            Id = id;
        }
    }

    public struct LargeXlsxStyle
    {
        internal uint Id { get; }

        internal LargeXlsxStyle(uint id)
        {
            Id = id;
        }
    }
}