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
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    /*
     * Special thanks to http://polymathprogrammer.com/2009/11/09/how-to-create-stylesheet-in-excel-open-xml/
     * for very valuable insights on how to properly create styles
     * and to https://github.com/ClosedXML/ClosedXML/wiki/NumberFormatId-Lookup-Table
     * for built-in number formats.
     */
    internal class Stylesheet
    {
        private readonly Dictionary<XlsxFont, int> _fonts;
        private readonly Dictionary<XlsxFill, int> _fills;
        private readonly Dictionary<XlsxBorder, int> _borders;
        private readonly Dictionary<XlsxNumberFormat, int> _numberFormats;
        private readonly Dictionary<XlsxStyle, int> _styles;
        private int _nextFontId;
        private int _nextBorderId;
        private int _nextFillId;
        private int _nextNumberFormatId;
        private int _nextStyleId;

        public Stylesheet()
        {
            _fonts = new Dictionary<XlsxFont, int> { [XlsxFont.Default] = 0 };
            _nextFontId = 1;

            _fills = new Dictionary<XlsxFill, int> { [XlsxFill.None] = 0, [XlsxFill.Gray125] = 1 };
            _nextFillId = 2; // ids less than 2 are hardcoded by Excel for default fills

            _borders = new Dictionary<XlsxBorder, int> { [XlsxBorder.None] = 0 };
            _nextBorderId = 1;

            _numberFormats = new Dictionary<XlsxNumberFormat, int>
            {
                [XlsxNumberFormat.General] = 0,
                [XlsxNumberFormat.Integer] = 1,
                [XlsxNumberFormat.TwoDecimal] = 2,
                [XlsxNumberFormat.ThousandInteger] = 3,
                [XlsxNumberFormat.ThousandTwoDecimal] = 4,
                [XlsxNumberFormat.IntegerPercentage] = 9,
                [XlsxNumberFormat.TwoDecimalPercentage] = 10,
                [XlsxNumberFormat.Scientific] = 11,
                [XlsxNumberFormat.ShortDate] = 14,
                [XlsxNumberFormat.ShortDateTime] = 22
            };
            _nextNumberFormatId = 164; // ids less than 164 are hardcoded by Excel for default formats

            _styles = new Dictionary<XlsxStyle, int> { [XlsxStyle.Default] = 0 };
            _nextStyleId = 1;
        }

        public int ResolveStyleId(XlsxStyle style)
        {
            if (_styles.TryGetValue(style, out var id))
                return id;

            if (!_fonts.TryGetValue(style.Font, out var fontId))
            {
                fontId = _nextFontId++;
                _fonts.Add(style.Font, fontId);
            }
            if (!_fills.TryGetValue(style.Fill, out var fillId))
            {
                fillId = _nextFillId++;
                _fills.Add(style.Fill, fillId);
            }
            if (!_borders.TryGetValue(style.Border, out var borderId))
            {
                borderId = _nextBorderId++;
                _borders.Add(style.Border, borderId);
            }
            if (!_numberFormats.TryGetValue(style.NumberFormat, out var numberFormatId))
            {
                numberFormatId = _nextNumberFormatId++;
                _numberFormats.Add(style.NumberFormat, numberFormatId);
            }
            id = _nextStyleId++;
            _styles.Add(style, id);
            return id;
        }

        public void Save(ZipWriter zipWriter)
        {
            using (var stream = zipWriter.WriteToStream("xl/styles.xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                WriteNumberFormats(streamWriter);
                WriteFonts(streamWriter);
                WriteFills(streamWriter);
                WriteBorders(streamWriter);
                WriteCellFormats(streamWriter);
                streamWriter.Write("</styleSheet>");
            }
        }

        private void WriteNumberFormats(StreamWriter streamWriter)
        {
            streamWriter.Write("<numFmts count=\"{0}\">", _numberFormats.Count);
            foreach (var numberFormat in _numberFormats.OrderBy(nf => nf.Value))
            {
                streamWriter.Write("<numFmt numFmtId=\"{0}\" formatCode=\"{1}\"/>",
                    numberFormat.Value, Util.EscapeXmlAttribute(numberFormat.Key.FormatCode));
            }
            streamWriter.Write("</numFmts>");
        }

        private void WriteFonts(StreamWriter streamWriter)
        {
            streamWriter.Write("<fonts count=\"{0}\">", _fonts.Count);
            foreach (var font in _fonts.OrderBy(f => f.Value))
            {
                streamWriter.Write("<font>"
                                   + "<sz val=\"{0}\"/>"
                                   + "<color rgb=\"{1}\"/>"
                                   + "<name val=\"{2}\"/>"
                                   + "<family val=\"2\"/>"
                                   + "{3}{4}{5}"
                                   + "</font>",
                    font.Key.FontSize, GetColorString(font.Key.Color), Util.EscapeXmlAttribute(font.Key.FontName),
                    font.Key.Bold ? "<b/>" : "", font.Key.Italic ? "<i/>" : "", font.Key.Strike ? "<strike/>" : "");
            }
            streamWriter.Write("</fonts>");
        }

        private void WriteFills(StreamWriter streamWriter)
        {
            streamWriter.Write("<fills count=\"{0}\">", _fills.Count);
            foreach (var fill in _fills.OrderBy(f => f.Value))
            {
                streamWriter.Write("<fill>"
                                   + "<patternFill patternType=\"{0}\">"
                                   + "<fgColor rgb=\"{1}\"/>"
                                   + "<bgColor rgb=\"{1}\"/>"
                                   + "</patternFill>"
                                   + "</fill>",
                    EnumToAttributeValue(fill.Key.PatternType), GetColorString(fill.Key.Color));
            }
            streamWriter.Write("</fills>");
        }

        private void WriteBorders(StreamWriter streamWriter)
        {
            streamWriter.Write($"<borders count=\"{_borders.Count}\">");
            foreach (var border in _borders.OrderBy(b => b.Value))
            {
                streamWriter.Write($"<border diagonalDown=\"{BoolToInt(border.Key.DiagonalDown)}\" diagonalUp=\"{BoolToInt(border.Key.DiagonalUp)}\">");
                WriteBorderLine(streamWriter, "left", border.Key.Left);
                WriteBorderLine(streamWriter, "right", border.Key.Right);
                WriteBorderLine(streamWriter, "top", border.Key.Top);
                WriteBorderLine(streamWriter, "bottom", border.Key.Bottom);
                WriteBorderLine(streamWriter, "diagonal", border.Key.Diagonal);
                streamWriter.Write("</border>");
            }
            streamWriter.Write("</borders>");
        }

        private static void WriteBorderLine(StreamWriter streamWriter, string elementName, XlsxBorder.Line line)
        {
            if (line != null)
            {
                streamWriter.Write($"<{elementName} style=\"{EnumToAttributeValue(line.Style)}\">");
                if (line.Color != Color.Transparent)
                    streamWriter.Write($"<color rgb=\"{GetColorString(line.Color)}\"/>");
                streamWriter.Write($"</{elementName}>");
            }
            else
            {
                streamWriter.Write($"<{elementName}/>");
            }
        }

        private void WriteCellFormats(StreamWriter streamWriter)
        {
            streamWriter.Write("<cellXfs count=\"{0}\">", _styles.Count);
            foreach (var style in _styles.OrderBy(s => s.Value))
            {
                streamWriter.Write("<xf numFmtId=\"{0}\" fontId=\"{1}\" fillId=\"{2}\" borderId=\"{3}\""
                                   + " applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"",
                    _numberFormats[style.Key.NumberFormat], _fonts[style.Key.Font], _fills[style.Key.Fill],
                    _borders[style.Key.Border]);
                if (style.Key.Alignment != null)
                {
                    streamWriter.Write(" applyAlignment=\"1\"><alignment");
                    var a = style.Key.Alignment;
                    if (a.HorizontalType != XlsxAlignment.Horizontal.General) streamWriter.Write(" horizontal=\"{0}\"", EnumToAttributeValue(a.HorizontalType));
                    if (a.VerticalType != XlsxAlignment.Vertical.Bottom) streamWriter.Write(" vertical=\"{0}\"", EnumToAttributeValue(a.VerticalType));
                    if (a.Indent != 0) streamWriter.Write(" indent=\"{0}\"", a.Indent);
                    if (a.JustifyLastLine) streamWriter.Write(" justifyLastLine=\"1\"");
                    if (a.ReadingOrderType != XlsxAlignment.ReadingOrder.ContextDependent) streamWriter.Write(" readingOrder=\"{0}\"", (int)a.ReadingOrderType);
                    if (a.ShrinkToFit) streamWriter.Write(" shrinkToFit=\"1\"");
                    if (a.TextRotation != 0) streamWriter.Write(" textRotation=\"{0}\"", a.TextRotation);
                    if (a.WrapText) streamWriter.Write(" wrapText=\"1\"");
                    streamWriter.Write("/></xf>");
                }
                else
                {
                    streamWriter.Write("/>");
                }
            }
            streamWriter.Write("</cellXfs>");
        }

        private static string GetColorString(Color color) => $"{color.R:x2}{color.G:x2}{color.B:x2}";

        private static int BoolToInt(bool value) => value ? 1 : 0;

        private static string EnumToAttributeValue<T>(T enumValue)
        {
            var s = enumValue.ToString();
            return char.ToLowerInvariant(s[0]) + s.Substring(1);
        }
    }
}