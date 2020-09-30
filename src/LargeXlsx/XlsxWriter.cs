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
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SharpCompress.Common;
using SharpCompress.Writers;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    public class XlsxWriter : IDisposable
    {
        private const int MaxSheetNameLength = 31;
        private readonly ZipWriter _zipWriter;
        private readonly List<Worksheet> _worksheets;
        private readonly Stylesheet _stylesheet;
        private Worksheet _currentWorksheet;
        private bool _hasFormulasWithoutResult;

        public XlsxStyle DefaultStyle { get; private set; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;
        public string CurrentColumnName => Util.GetColumnName(CurrentColumnNumber);
        public string GetRelativeColumnName(int offsetFromCurrentColumn) => Util.GetColumnName(CurrentColumnNumber + offsetFromCurrentColumn);
        public static string GetColumnName(int columnIndex) => Util.GetColumnName(columnIndex);

        public XlsxWriter(Stream stream)
        {
            _worksheets = new List<Worksheet>();
            _stylesheet = new Stylesheet();
            DefaultStyle = XlsxStyle.Default;

            _zipWriter = (ZipWriter)WriterFactory.Open(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate));
        }

        public void Dispose()
        {
            _currentWorksheet?.Dispose();
            _stylesheet.Save(_zipWriter);
            Save();
            _zipWriter.Dispose();
        }

        private void Save()
        {
            using (var stream = _zipWriter.WriteToStream("[Content_Types].xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                var worksheetTags = new StringBuilder();
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<Override PartName=\"/xl/worksheets/sheet{worksheet.Id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                                   + "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                                   + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + worksheetTags
                                   + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                                   + "</Types>");
            }

            using (var stream = _zipWriter.WriteToStream("_rels/.rels", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"RidWB1\"/>"
                                   + "</Relationships>");
            }

            using (var stream = _zipWriter.WriteToStream("xl/workbook.xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                var worksheetTags = new StringBuilder();
                var definedNames = new StringBuilder();
                var sheetIndex = 0;
                foreach (var worksheet in _worksheets)
                {
                    worksheetTags.Append($"<sheet name=\"{Util.EscapeXmlAttribute(worksheet.Name)}\" sheetId=\"{worksheet.Id}\" r:id=\"RidWS{worksheet.Id}\"/>");
                    if (worksheet.AutoFilterAbsoluteRef != null)
                        definedNames.Append($"<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"{sheetIndex}\" hidden=\"1\">{Util.EscapeXmlText(worksheet.AutoFilterAbsoluteRef)}</definedName>");
                    sheetIndex++;
                }
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                                   + "<sheets>"
                                   + worksheetTags
                                   + "</sheets>");
                if (definedNames.Length > 0) streamWriter.Write("<definedNames>" + definedNames + "</definedNames>");
                if (_hasFormulasWithoutResult) streamWriter.Write("<calcPr calcCompleted=\"0\" fullCalcOnLoad=\"1\"/>");
                streamWriter.Write("</workbook>");
            }

            using (var stream = _zipWriter.WriteToStream("xl/_rels/workbook.xml.rels", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                var worksheetTags = new StringBuilder();
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet{worksheet.Id}.xml\" Id=\"RidWS{worksheet.Id}\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + worksheetTags
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"RidSS1\"/>"
                                   + "</Relationships>");
            }
        }

        public XlsxWriter BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0, IEnumerable<XlsxColumn> columns = null)
        {
            if (name.Length > MaxSheetNameLength)
                throw new ArgumentException($"The name \"{name}\" exceeds the maximum length of {MaxSheetNameLength} characters supported by Excel");
            if (_worksheets.Any(ws => string.Equals(ws.Name, name, StringComparison.InvariantCultureIgnoreCase)))
                throw new ArgumentException($"A worksheet named \"{name}\" has already been added");
            _currentWorksheet?.Dispose();
            _currentWorksheet = new Worksheet(_zipWriter, _worksheets.Count + 1, name, splitRow, splitColumn, _stylesheet, columns ?? Enumerable.Empty<XlsxColumn>());
            _worksheets.Add(_currentWorksheet);
            return this;
        }

        public XlsxWriter SkipRows(int rowCount)
        {
            return DoOnWorksheet(() => _currentWorksheet.SkipRows(rowCount));
        }

        public XlsxWriter BeginRow(double? height = null, bool hidden = false, XlsxStyle style = null)
        {
            return DoOnWorksheet(() => _currentWorksheet.BeginRow(height, hidden, style));
        }

        public XlsxWriter SkipColumns(int columnCount)
        {
            return DoOnWorksheet(() => _currentWorksheet.SkipColumns(columnCount));
        }

        public XlsxWriter Write(XlsxStyle style = null, int columnSpan = 1, int repeatCount = 1)
        {
            if (columnSpan == 1)
                return DoOnWorksheet(() => _currentWorksheet.Write(style ?? DefaultStyle, repeatCount));
            
            for (var i = 0; i < repeatCount; i++)
                AddMergedCell(1, columnSpan).Write(style, 1).Write(style, repeatCount: columnSpan - 1);
            return this;
        }

        public XlsxWriter Write(string value, XlsxStyle style = null, int columnSpan = 1)
        {
            return columnSpan == 1
                ? DoOnWorksheet(() => _currentWorksheet.Write(value, style ?? DefaultStyle))
                : AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter Write(double value, XlsxStyle style = null, int columnSpan = 1)
        {
            return columnSpan == 1
                ? DoOnWorksheet(() => _currentWorksheet.Write(value, style ?? DefaultStyle))
                : AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter Write(decimal value, XlsxStyle style = null, int columnSpan = 1)
        {
            return Write((double)value, style, columnSpan);
        }

        public XlsxWriter Write(int value, XlsxStyle style = null, int columnSpan = 1)
        {
            return Write((double)value, style, columnSpan);
        }

        public XlsxWriter Write(DateTime value, XlsxStyle style = null, int columnSpan = 1)
        {
            return Write(Util.DateToDouble(value), style, columnSpan);
        }

        public XlsxWriter WriteFormula(string formula, XlsxStyle style = null, int columnSpan = 1, IConvertible result = null)
        {
            return columnSpan == 1
                ? DoOnWorksheet(() =>
                {
                    if (result == null) _hasFormulasWithoutResult = true;
                    _currentWorksheet.WriteFormula(formula, style ?? DefaultStyle, result);
                })
                : AddMergedCell(1, columnSpan).WriteFormula(formula, style, 1, result).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            return DoOnWorksheet(() => _currentWorksheet.AddMergedCell(fromRow, fromColumn, rowCount, columnCount));
        }

        public XlsxWriter AddMergedCell(int rowCount, int columnCount)
        {
            return AddMergedCell(CurrentRowNumber, CurrentColumnNumber, rowCount, columnCount);
        }

        public XlsxWriter SetAutoFilter(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            return DoOnWorksheet(() => _currentWorksheet.SetAutoFilter(fromRow, fromColumn, rowCount, columnCount));
        }

        public XlsxWriter AddDataValidation(int fromRow, int fromColumn, int rowCount, int columnCount, XlsxDataValidation dataValidation)
        {
            return DoOnWorksheet(() => _currentWorksheet.AddDataValidation(fromRow, fromColumn, rowCount, columnCount, dataValidation));
        }

        public XlsxWriter AddDataValidation(int rowCount, int columnCount, XlsxDataValidation dataValidation)
        {
            return AddDataValidation(CurrentRowNumber, CurrentColumnNumber, rowCount, columnCount, dataValidation);
        }

        public XlsxWriter AddDataValidation(XlsxDataValidation dataValidation)
        {
            return AddDataValidation(CurrentRowNumber, CurrentColumnNumber, 1, 1, dataValidation);
        }

        public XlsxWriter SetDefaultStyle(XlsxStyle style)
        {
            DefaultStyle = style;
            return this;
        }

        private XlsxWriter DoOnWorksheet(Action action)
        {
            if (_currentWorksheet == null)
                throw new InvalidOperationException($"{nameof(BeginWorksheet)} not called");
            action();
            return this;
        }
    }
}
