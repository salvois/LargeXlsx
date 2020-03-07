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
using System.IO;
using System.Text;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    /*
     * Special thanks to http://polymathprogrammer.com/2009/11/09/how-to-create-stylesheet-in-excel-open-xml/
     * for very valuable insights on how to properly create styles.
     */
    public class XlsxStylesheet
    {
        private readonly List<XlsxFont> _fonts;
        private readonly List<XlsxFill> _fills;
        private readonly List<XlsxBorder> _borders;
        private readonly List<XlsxNumberFormat> _numberFormats;
        private readonly List<XlsxStyle> _styles;
        private readonly Dictionary<StyleTuple, XlsxStyle> _deduplicatedStyles;
        private int _nextFontId;
        private int _nextBorderId;
        private int _nextFillId;
        private int _nextNumberFormatId;
        private int _nextStyleId;

        internal XlsxStylesheet()
        {
            _fonts = new List<XlsxFont>();
            _fills = new List<XlsxFill>();
            _borders = new List<XlsxBorder>();
            _numberFormats = new List<XlsxNumberFormat>();
            _styles = new List<XlsxStyle>();
            _deduplicatedStyles = new Dictionary<StyleTuple, XlsxStyle>();

            _fonts.Add(XlsxFont.Default);
            _fills.Add(XlsxFill.None);
            _fills.Add(XlsxFill.Gray125);
            _borders.Add(XlsxBorder.None);
            _styles.Add(XlsxStyle.Default);
            _deduplicatedStyles.Add(new StyleTuple(XlsxStyle.Default.Font.Id, XlsxStyle.Default.Fill.Id, XlsxStyle.Default.NumberFormat.Id, XlsxStyle.Default.Border.Id),
                                    XlsxStyle.Default);

            _nextFillId = XlsxFill.FirstAvailableId;
            _nextBorderId = XlsxBorder.FirstAvailableId;
            _nextNumberFormatId = XlsxNumberFormat.FirstAvailableId;
            _nextFontId = XlsxFont.FirstAvailableId;
            _nextStyleId = XlsxStyle.FirstAvailableId;
        }

        public XlsxFont CreateFont(string fontName, double fontSize, string hexRgbColor, bool bold = false, bool italic = false, bool strike = false)
        {
            var font = new XlsxFont(_nextFontId++, fontName, fontSize, hexRgbColor, bold, italic, strike);
            _fonts.Add(font);
            return font;
        }

        public XlsxFill CreateSolidFill(string hexRgbColor)
        {
            var fill = new XlsxFill(_nextFillId++, XlsxFill.Pattern.Solid, hexRgbColor);
            _fills.Add(fill);
            return fill;
        }

        public XlsxNumberFormat CreateNumberFormat(string formatCode)
        {
            var numberFormat = new XlsxNumberFormat(_nextNumberFormatId++, formatCode);
            _numberFormats.Add(numberFormat);
            return numberFormat;
        }

        public XlsxBorder CreateBorder(
            string hexRgbColor,
            XlsxBorder.Style top = XlsxBorder.Style.None,
            XlsxBorder.Style right = XlsxBorder.Style.None,
            XlsxBorder.Style bottom = XlsxBorder.Style.None,
            XlsxBorder.Style left = XlsxBorder.Style.None)
        {
            var border = new XlsxBorder(_nextBorderId++, hexRgbColor, top, right, bottom, left);
            _borders.Add(border);
            return border;
        }

        public XlsxStyle CreateStyle(XlsxFont font, XlsxFill fill, XlsxBorder border, XlsxNumberFormat numberFormat)
        {
            var styleTuple = new StyleTuple(font.Id, fill.Id, numberFormat.Id, border.Id);
            if (_deduplicatedStyles.TryGetValue(styleTuple, out var style))
                return style;

            var newStyle = new XlsxStyle(_nextStyleId++, font, fill, border, numberFormat);
            _styles.Add(newStyle);
            _deduplicatedStyles[styleTuple] = newStyle;
            return newStyle;
        }

        internal void Save(ZipWriter zipWriter)
        {
            using (var stream = zipWriter.WriteToStream("xl/styles.xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");

                streamWriter.Write("<numFmts count=\"{0}\">", _numberFormats.Count);
                foreach (var numberFormat in _numberFormats)
                {
                    streamWriter.Write("<numFmt numFmtId=\"{0}\" formatCode=\"{1}\"/>", numberFormat.Id, numberFormat.FormatCode);
                }
                streamWriter.Write("</numFmts>");

                streamWriter.Write("<fonts count=\"{0}\">", _fonts.Count);
                foreach (var font in _fonts)
                {
                    streamWriter.Write("<font>"
                                       + "<sz val=\"{0}\"/>"
                                       + "<color rgb=\"{1}\"/>"
                                       + "<name val=\"{2}\"/>"
                                       + "<family val=\"2\"/>"
                                       + "{3}{4}{5}"
                                       + "</font>",
                        font.FontSize, font.HexRgbColor, font.FontName,
                        font.Bold ? "<b/>" : "", font.Italic ? "<i/>" : "", font.Strike ? "<strike/>" : "");
                }
                streamWriter.Write("</fonts>");

                streamWriter.Write("<fills count=\"{0}\">", _fills.Count);
                foreach (var fill in _fills)
                {
                    streamWriter.Write("<fill>"
                                       + "<patternFill patternType=\"{0}\">"
                                       + "<fgColor rgb=\"{1}\"/>"
                                       + "<bgColor rgb=\"{1}\"/>"
                                       + "</patternFill>"
                                       + "</fill>",
                        XlsxFill.GetPatternAttributeValue(fill.PatternType), fill.HexRgbColor);
                }
                streamWriter.Write("</fills>");

                streamWriter.Write("<borders count=\"{0}\">", _borders.Count);
                foreach (var border in _borders)
                {
                    streamWriter.Write("<border>"
                                       + "<left color=\"{0}\" style=\"{4}\"/>"
                                       + "<right color=\"{0}\" style=\"{2}\"/>"
                                       + "<top color=\"{0}\" style=\"{1}\"/>"
                                       + "<bottom color=\"{0}\" style=\"{3}\"/>"
                                       + "<diagonal/>"
                                       + "</border>",
                        border.HexRgbColor,
                        XlsxBorder.GetStyleAttributeValue(border.Top),
                        XlsxBorder.GetStyleAttributeValue(border.Right),
                        XlsxBorder.GetStyleAttributeValue(border.Bottom),
                        XlsxBorder.GetStyleAttributeValue(border.Left));
                }
                streamWriter.Write("</borders>");

                streamWriter.Write("<cellXfs count=\"{0}\">", _styles.Count);
                foreach (var style in _styles)
                {
                    streamWriter.Write("<xf numFmtId=\"{0}\" fontId=\"{1}\" fillId=\"{2}\" borderId=\"{3}\""
                                       + " applyNumberFormat=\"1\" applyFont=\"1\" applyFill=\"1\" applyBorder=\"1\"/>",
                        style.NumberFormat.Id, style.Font.Id, style.Fill.Id, style.Border.Id);
                }
                streamWriter.Write("</cellXfs>");

                streamWriter.Write("</styleSheet>");
            }
        }

        private struct StyleTuple
        {
            public int FontId;
            public int FillId;
            public int NumberFormatId;
            public int BorderId;

            public StyleTuple(int fontId, int fillId, int numberFormatId, int borderId)
            {
                FontId = fontId;
                FillId = fillId;
                NumberFormatId = numberFormatId;
                BorderId = borderId;
            }
        }
    }
}