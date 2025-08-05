/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025 Salvatore ISAJA. All rights reserved.

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
using System.IO.Compression;
using System.Linq;

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
        private const int FirstCustomNumberFormatId = 164; // ids less than 164 are hardcoded by Excel for default formats
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
        private XlsxStyle _lastUsedStyle;
        private int _lastUsedStyleId;

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
                [XlsxNumberFormat.ShortDateTime] = 22,
                [XlsxNumberFormat.Text] = 49
            };
            _nextNumberFormatId = FirstCustomNumberFormatId;

            _styles = new Dictionary<XlsxStyle, int> { [XlsxStyle.Default] = 0 };
            _nextStyleId = 1;
            _lastUsedStyle = XlsxStyle.Default;
            _lastUsedStyleId = 0;
        }

        public int ResolveStyleId(XlsxStyle style)
        {
            if (ReferenceEquals(style, _lastUsedStyle))
                return _lastUsedStyleId;
            if (!_styles.TryGetValue(style, out var id))
            {
                if (!_fonts.ContainsKey(style.Font))
                    _fonts.Add(style.Font, _nextFontId++);
                if (!_fills.ContainsKey(style.Fill))
                    _fills.Add(style.Fill, _nextFillId++);
                if (!_borders.ContainsKey(style.Border))
                    _borders.Add(style.Border, _nextBorderId++);
                if (!_numberFormats.ContainsKey(style.NumberFormat))
                    _numberFormats.Add(style.NumberFormat, _nextNumberFormatId++);
                id = _nextStyleId++;
                _styles.Add(style, id);
            }
            SetLastUsedStyle(style, id);
            return id;
        }

        public void Save(ZipArchive zipArchive)
        {
            var entry = zipArchive.CreateEntry("xl/styles.xml", CompressionLevel.Optimal);
            using (var streamWriter = new InvariantCultureStreamWriter(entry.Open()))
            {
                streamWriter.WriteLine("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
                WriteNumberFormats(streamWriter);
                WriteFonts(streamWriter);
                WriteFills(streamWriter);
                WriteBorders(streamWriter);
                WriteCellFormats(streamWriter);
                streamWriter.WriteLine("</styleSheet>");
            }
        }

        private void SetLastUsedStyle(XlsxStyle style, int styleId)
        {
            _lastUsedStyle = style;
            _lastUsedStyleId = styleId;
        }

        private void WriteNumberFormats(StreamWriter streamWriter)
        {
            streamWriter.WriteLine("<numFmts count=\"{0}\">", _numberFormats.Count(nf => nf.Value >= FirstCustomNumberFormatId));
            foreach (var numberFormat in _numberFormats.Where(nf => nf.Value >= FirstCustomNumberFormatId).OrderBy(nf => nf.Value))
            {
                streamWriter
                    .Append("<numFmt numFmtId=\"")
                    .Append(numberFormat.Value)
                    .Append("\" formatCode=\"")
                    .AppendEscapedXmlAttribute(numberFormat.Key.FormatCode, false)
                    .Append("\"/>\n");
            }
            streamWriter.WriteLine("</numFmts>");
        }

        private void WriteFonts(StreamWriter streamWriter)
        {
            streamWriter.WriteLine("<fonts count=\"{0}\">", _fonts.Count);
            foreach (var font in _fonts.OrderBy(f => f.Value))
            {
                streamWriter
                    .Append("<font><sz val=\"")
                    .Append(font.Key.Size)
                    .Append("\"/><color rgb=\"")
                    .Append(GetColorString(font.Key.Color))
                    .Append("\"/><name val=\"")
                    .AppendEscapedXmlAttribute(font.Key.Name, false)
                    .Append("\"/><family val=\"2\"/>");
                if (font.Key.Bold)
                    streamWriter.Append("<b val=\"true\"/>");
                if (font.Key.Italic)
                    streamWriter.Append("<i val=\"true\"/>");
                if (font.Key.Strike)
                    streamWriter.Append("<strike val=\"true\"/>");
                switch (font.Key.UnderlineType)
                {
                    case XlsxFont.Underline.None:
                        break;
                    case XlsxFont.Underline.Single:
                        streamWriter.Append("<u/>");
                        break;
                    default:
                        streamWriter.Append($"<u val=\"{Util.EnumToAttributeValue(font.Key.UnderlineType)}\"/>");
                        break;
                }
                streamWriter.Append("</font>\n");
            }
            streamWriter.WriteLine("</fonts>");
        }

        private void WriteFills(StreamWriter streamWriter)
        {
            streamWriter.WriteLine("<fills count=\"{0}\">", _fills.Count);
            foreach (var fill in _fills.OrderBy(f => f.Value))
            {
                streamWriter.WriteLine("<fill>"
                                   + "<patternFill patternType=\"{0}\">"
                                   + "<fgColor rgb=\"{1}\"/>"
                                   + "<bgColor rgb=\"{1}\"/>"
                                   + "</patternFill>"
                                   + "</fill>",
                    Util.EnumToAttributeValue(fill.Key.PatternType), GetColorString(fill.Key.Color));
            }
            streamWriter.WriteLine("</fills>");
        }

        private void WriteBorders(StreamWriter streamWriter)
        {
            streamWriter.WriteLine($"<borders count=\"{_borders.Count}\">");
            foreach (var border in _borders.OrderBy(b => b.Value))
            {
                streamWriter.WriteLine($"<border diagonalDown=\"{Util.BoolToInt(border.Key.DiagonalDown)}\" diagonalUp=\"{Util.BoolToInt(border.Key.DiagonalUp)}\">");
                WriteBorderLine(streamWriter, "left", border.Key.Left);
                WriteBorderLine(streamWriter, "right", border.Key.Right);
                WriteBorderLine(streamWriter, "top", border.Key.Top);
                WriteBorderLine(streamWriter, "bottom", border.Key.Bottom);
                WriteBorderLine(streamWriter, "diagonal", border.Key.Diagonal);
                streamWriter.WriteLine("</border>");
            }
            streamWriter.WriteLine("</borders>");
        }

        private static void WriteBorderLine(StreamWriter streamWriter, string elementName, XlsxBorder.Line line)
        {
            if (line != null)
            {
                streamWriter.Write($"<{elementName} style=\"{Util.EnumToAttributeValue(line.Style)}\">");
                if (line.Color != Color.Transparent)
                    streamWriter.Write($"<color rgb=\"{GetColorString(line.Color)}\"/>");
                streamWriter.WriteLine($"</{elementName}>");
            }
            else
            {
                streamWriter.WriteLine($"<{elementName}/>");
            }
        }

        private void WriteCellFormats(StreamWriter streamWriter)
        {
            streamWriter.WriteLine("<cellXfs count=\"{0}\">", _styles.Count);
            foreach (var style in _styles.OrderBy(s => s.Value))
            {
                streamWriter.Write("<xf numFmtId=\"{0}\" fontId=\"{1}\" fillId=\"{2}\" borderId=\"{3}\""
                                   + " applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"",
                    _numberFormats[style.Key.NumberFormat], _fonts[style.Key.Font], _fills[style.Key.Fill],
                    _borders[style.Key.Border]);
                if (style.Key.Alignment != XlsxAlignment.Default)
                {
                    streamWriter.Write(" applyAlignment=\"1\"><alignment");
                    var a = style.Key.Alignment;
                    if (a.HorizontalType != XlsxAlignment.Horizontal.General) streamWriter.Write(" horizontal=\"{0}\"", Util.EnumToAttributeValue(a.HorizontalType));
                    if (a.VerticalType != XlsxAlignment.Vertical.Bottom) streamWriter.Write(" vertical=\"{0}\"", Util.EnumToAttributeValue(a.VerticalType));
                    if (a.Indent != 0) streamWriter.Write(" indent=\"{0}\"", a.Indent);
                    if (a.JustifyLastLine) streamWriter.Write(" justifyLastLine=\"1\"");
                    if (a.ReadingOrderType != XlsxAlignment.ReadingOrder.ContextDependent) streamWriter.Write(" readingOrder=\"{0}\"", (int)a.ReadingOrderType);
                    if (a.ShrinkToFit) streamWriter.Write(" shrinkToFit=\"1\"");
                    if (a.TextRotation != 0) streamWriter.Write(" textRotation=\"{0}\"", a.TextRotation);
                    if (a.WrapText) streamWriter.Write(" wrapText=\"1\"");
                    streamWriter.WriteLine("/></xf>");
                }
                else
                {
                    streamWriter.WriteLine("/>");
                }
            }
            streamWriter.WriteLine("</cellXfs>");
        }

        private static string GetColorString(Color color) => $"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}