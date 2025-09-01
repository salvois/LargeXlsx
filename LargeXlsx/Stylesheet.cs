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
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
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

        public void Save(ZipWriter zipWriter, CustomWriter customWriter)
        {
            using (var stream = zipWriter.WriteToStream("xl/styles.xml", new ZipWriterEntryOptions()))
            {
                customWriter.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n"u8
                                   + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">\n"u8);
                WriteNumberFormats(customWriter);
                WriteFonts(customWriter);
                WriteFills(customWriter);
                WriteBorders(customWriter);
                WriteCellFormats(customWriter);
                customWriter.Append("</styleSheet>\n"u8);
                customWriter.FlushTo(stream);
            }
        }

        private void SetLastUsedStyle(XlsxStyle style, int styleId)
        {
            _lastUsedStyle = style;
            _lastUsedStyleId = styleId;
        }

        private void WriteNumberFormats(CustomWriter customWriter)
        {
            customWriter.Append("<numFmts count=\""u8).Append(_numberFormats.Count(nf => nf.Value >= FirstCustomNumberFormatId)).Append("\">\n"u8);
            foreach (var numberFormat in _numberFormats.Where(nf => nf.Value >= FirstCustomNumberFormatId).OrderBy(nf => nf.Value))
            {
                customWriter
                    .Append("<numFmt numFmtId=\""u8)
                    .Append(numberFormat.Value)
                    .Append("\" formatCode=\""u8)
                    .AppendEscapedXmlAttribute(numberFormat.Key.FormatCode, false)
                    .Append("\"/>\n"u8);
            }
            customWriter.Append("</numFmts>\n"u8);
        }

        private void WriteFonts(CustomWriter customWriter)
        {
            customWriter.Append("<fonts count=\""u8).Append(_fonts.Count).Append("\">\n"u8);
            foreach (var font in _fonts.OrderBy(f => f.Value))
            {
                customWriter
                    .Append("<font><sz val=\""u8)
                    .Append(font.Key.Size)
                    .Append("\"/><color rgb=\""u8)
                    .AppendEscapedXmlAttribute(GetColorString(font.Key.Color), false)
                    .Append("\"/><name val=\""u8)
                    .AppendEscapedXmlAttribute(font.Key.Name, false)
                    .Append("\"/><family val=\"2\"/>"u8);
                if (font.Key.Bold)
                    customWriter.Append("<b val=\"true\"/>"u8);
                if (font.Key.Italic)
                    customWriter.Append("<i val=\"true\"/>"u8);
                if (font.Key.Strike)
                    customWriter.Append("<strike val=\"true\"/>"u8);
                switch (font.Key.UnderlineType)
                {
                    case XlsxFont.Underline.None:
                        break;
                    case XlsxFont.Underline.Single:
                        customWriter.Append("<u/>"u8);
                        break;
                    default:
                        customWriter.Append("<u val=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(font.Key.UnderlineType), false).Append("\"/>"u8);
                        break;
                }
                customWriter.Append("</font>\n"u8);
            }
            customWriter.Append("</fonts>\n"u8);
        }

        private void WriteFills(CustomWriter customWriter)
        {
            customWriter.Append("<fills count=\""u8).Append(_fills.Count).Append("\">\n"u8);
            foreach (var fill in _fills.OrderBy(f => f.Value))
            {
                var colorString = GetColorString(fill.Key.Color);
                customWriter.Append("<fill><patternFill patternType=\""u8)
                    .AppendEscapedXmlAttribute(Util.EnumToAttributeValue(fill.Key.PatternType), false)
                    .Append("\"><fgColor rgb=\""u8)
                    .AppendEscapedXmlAttribute(colorString, false)
                    .Append("\"/><bgColor rgb=\""u8)
                    .AppendEscapedXmlAttribute(colorString, false)
                    .Append("\"/></patternFill></fill>\n"u8);
            }
            customWriter.Append("</fills>\n"u8);
        }

        private void WriteBorders(CustomWriter customWriter)
        {
            customWriter.Append("<borders count=\""u8).Append(_borders.Count).Append("\">\n"u8);
            foreach (var border in _borders.OrderBy(b => b.Value))
            {
                customWriter
                    .Append("<border diagonalDown=\""u8)
                    .Append(Util.BoolToInt(border.Key.DiagonalDown))
                    .Append("\" diagonalUp=\""u8)
                    .Append(Util.BoolToInt(border.Key.DiagonalUp))
                    .Append("\">\n"u8);
                WriteBorderLine(customWriter, "left"u8, border.Key.Left);
                WriteBorderLine(customWriter, "right"u8, border.Key.Right);
                WriteBorderLine(customWriter, "top"u8, border.Key.Top);
                WriteBorderLine(customWriter, "bottom"u8, border.Key.Bottom);
                WriteBorderLine(customWriter, "diagonal"u8, border.Key.Diagonal);
                customWriter.Append("</border>\n"u8);
            }
            customWriter.Append("</borders>\n"u8);
        }

        private static void WriteBorderLine(CustomWriter customWriter, ReadOnlySpan<byte> elementName, XlsxBorder.Line line)
        {
            if (line != null)
            {
                customWriter.Append("<"u8).Append(elementName).Append(" style=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(line.Style), false).Append("\">"u8);
                if (line.Color != Color.Transparent)
                    customWriter.Append("<color rgb=\""u8).AppendEscapedXmlAttribute(GetColorString(line.Color), false).Append("\"/>"u8);
                customWriter.Append("</"u8).Append(elementName).Append(">\n"u8);
            }
            else
            {
                customWriter.Append("<"u8).Append(elementName).Append("/>\n"u8);
            }
        }

        private void WriteCellFormats(CustomWriter customWriter)
        {
            customWriter.Append("<cellXfs count=\""u8).Append(_styles.Count).Append("\">\n"u8);
            foreach (var style in _styles.OrderBy(s => s.Value))
            {
                customWriter
                    .Append("<xf numFmtId=\""u8)
                    .Append(_numberFormats[style.Key.NumberFormat])
                    .Append("\" fontId=\""u8)
                    .Append(_fonts[style.Key.Font])
                    .Append("\" fillId=\""u8)
                    .Append(_fills[style.Key.Fill])
                    .Append("\" borderId=\""u8)
                    .Append(_borders[style.Key.Border])
                    .Append("\" applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\""u8);
                if (style.Key.Alignment != XlsxAlignment.Default)
                {
                    customWriter.Append(" applyAlignment=\"1\"><alignment"u8);
                    var a = style.Key.Alignment;
                    if (a.HorizontalType != XlsxAlignment.Horizontal.General) customWriter.Append(" horizontal=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(a.HorizontalType), false).Append("\""u8);
                    if (a.VerticalType != XlsxAlignment.Vertical.Bottom) customWriter.Append(" vertical=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(a.VerticalType), false).Append("\""u8);
                    if (a.Indent != 0) customWriter.Append(" indent=\""u8).Append(a.Indent).Append("\""u8);
                    if (a.JustifyLastLine) customWriter.Append(" justifyLastLine=\"1\""u8);
                    if (a.ReadingOrderType != XlsxAlignment.ReadingOrder.ContextDependent) customWriter.Append(" readingOrder=\""u8).Append((int)a.ReadingOrderType).Append("\""u8);
                    if (a.ShrinkToFit) customWriter.Append(" shrinkToFit=\"1\""u8);
                    if (a.TextRotation != 0) customWriter.Append(" textRotation=\""u8).Append(a.TextRotation).Append("\""u8);
                    if (a.WrapText) customWriter.Append(" wrapText=\"1\""u8);
                    customWriter.Append("/></xf>\n"u8);
                }
                else
                {
                    customWriter.Append("/>\n"u8);
                }
            }
            customWriter.Append("</cellXfs>\n"u8);
        }

        private static string GetColorString(Color color) => $"{color.A:X2}{color.R:X2}{color.G:X2}{color.B:X2}";
    }
}