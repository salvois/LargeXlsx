/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2019 Salvatore ISAJA. All rights reserved.

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
        public const uint NoFillId = 0;
        public const uint Gray125FillId = 1;
        public const uint DefaultFontId = 0;
        public const uint GeneralNumberFormatId = 0;
        public const uint TwoDecimalExcelNumberFormatId = 4; // #,##0.00
        public const uint NoBorderId = 0;
        public const uint DefaultStyleId = 0;

        private readonly Stylesheet _stylesheet;
        private readonly Dictionary<StyleTuple, uint> _styles;
        private uint _nextFontId;
        private uint _nextBorderId;
        private uint _nextFillId;
        private uint _nextNumberFormatId;
        private uint _nextStyleId;

        public LargeXlsxStylesheet()
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

            _styles = new Dictionary<StyleTuple, uint>();
            _nextStyleId = 0;
            CreateStyle(DefaultFontId, NoFillId, GeneralNumberFormatId, NoBorderId);
        }

        public uint CreateFont(string fontName, double fontSize, string hexRgbColor)
        {
            _stylesheet.Fonts.AppendChild(new Font
            {
                FontSize = new FontSize { Val = fontSize },
                Color = new Color { Rgb = HexBinaryValue.FromString(hexRgbColor) },
                FontName = new FontName { Val = fontName },
                FontFamilyNumbering = new FontFamilyNumbering { Val = 2 },
                FontScheme = new FontScheme { Val = FontSchemeValues.Minor }
            });
            return _nextFontId++;
        }

        public uint CreateSolidFill(string hexRgbColor)
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
            return _nextFillId++;
        }

        public uint CreateNumberFormat(string formatCode)
        {
            _stylesheet.NumberingFormats.AppendChild(new NumberingFormat
            {
                NumberFormatId = _nextNumberFormatId,
                FormatCode = formatCode
            });
            return _nextNumberFormatId++;
        }

        public uint CreateBorder(BorderStyleValues top, BorderStyleValues right, BorderStyleValues bottom, BorderStyleValues left, string hexRgbColor)
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
            return _nextBorderId++;
        }

        public uint CreateStyle(uint fontId, uint fillId, uint numberFormatId, uint borderId)
        {
            var styleTuple = new StyleTuple(fontId, fillId, numberFormatId, borderId);
            if (_styles.TryGetValue(styleTuple, out var styleId))
                return styleId;

            _stylesheet.CellFormats.AppendChild(new CellFormat
            {
                NumberFormatId = numberFormatId,
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                ApplyFont = true,
                ApplyFill = true,
                ApplyNumberFormat = true,
                ApplyBorder = true
            });
            _styles[styleTuple] = _nextStyleId;
            return _nextStyleId++;
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