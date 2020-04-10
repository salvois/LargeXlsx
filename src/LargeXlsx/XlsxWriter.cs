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
using System.Text;
using SharpCompress.Common;
using SharpCompress.Writers;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    public class XlsxWriter : IDisposable
    {
        private readonly ZipWriter _zipWriter;
        private readonly List<Worksheet> _worksheets;
        private readonly Stylesheet _stylesheet;
        private Worksheet _currentWorksheet;

        public XlsxStyle DefaultStyle { get; private set; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;

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
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<sheet name=\"{Util.EscapeXmlAttribute(worksheet.Name)}\" sheetId=\"{worksheet.Id}\" r:id=\"RidWS{worksheet.Id}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                   + "<sheets>"
                                   + worksheetTags
                                   + "</sheets>"
                                   + "</workbook>");
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

        public XlsxWriter BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0)
        {
            _currentWorksheet?.Dispose();
            _currentWorksheet = new Worksheet(_zipWriter, _worksheets.Count + 1, name, splitRow, splitColumn);
            _worksheets.Add(_currentWorksheet);
            return this;
        }

        public XlsxWriter SkipRows(int rowCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipRows(rowCount);
            return this;
        }

        public XlsxWriter BeginRow()
        {
            EnsureWorksheet();
            _currentWorksheet.BeginRow();
            return this;
        }

        public XlsxWriter SkipColumns(int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipColumns(columnCount);
            return this;
        }

        public XlsxWriter Write()
        {
            return Write(DefaultStyle);
        }

        public XlsxWriter Write(XlsxStyle style)
        {
            EnsureWorksheet();
            var styleId = _stylesheet.ResolveStyleId(style);
            _currentWorksheet.Write(styleId);
            return this;
        }

        public XlsxWriter Write(string value)
        {
            return Write(value, DefaultStyle);
        }

        public XlsxWriter Write(string value, XlsxStyle style)
        {
            EnsureWorksheet();
            var styleId = _stylesheet.ResolveStyleId(style);
            _currentWorksheet.Write(value, styleId);
            return this;
        }

        public XlsxWriter Write(double value)
        {
            return Write(value, DefaultStyle);
        }

        public XlsxWriter Write(double value, XlsxStyle style)
        {
            EnsureWorksheet();
            var styleId = _stylesheet.ResolveStyleId(style);
            _currentWorksheet.Write(value, styleId);
            return this;
        }

        public XlsxWriter Write(decimal value)
        {
            return Write(value, DefaultStyle);
        }

        public XlsxWriter Write(decimal value, XlsxStyle style)
        {
            EnsureWorksheet();
            var styleId = _stylesheet.ResolveStyleId(style);
            _currentWorksheet.Write((double)value, styleId);
            return this;
        }

        public XlsxWriter Write(int value)
        {
            return Write(value, DefaultStyle);
        }

        public XlsxWriter Write(int value, XlsxStyle style)
        {
            EnsureWorksheet();
            var styleId = _stylesheet.ResolveStyleId(style);
            _currentWorksheet.Write(value, styleId);
            return this;
        }

        public XlsxWriter AddMergedCell(int rowCount, int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.AddMergedCell(_currentWorksheet.CurrentRowNumber, _currentWorksheet.CurrentColumnNumber, rowCount, columnCount);
            return this;
        }

        public XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.AddMergedCell(fromRow, fromColumn, rowCount, columnCount);
            return this;
        }

        public XlsxWriter SetDefaultStyle(XlsxStyle style)
        {
            DefaultStyle = style;
            return this;
        }

        private void EnsureWorksheet()
        {
            if (_currentWorksheet == null)
                throw new InvalidOperationException($"{nameof(BeginWorksheet)} not called");
        }
    }
}
