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
    public class XlsxWriter2 : IDisposable
    {
        private readonly ZipWriter _zipWriter;
        private readonly List<XlsxWorksheet2> _worksheets;
        private XlsxWorksheet2 _currentWorksheet;

        public XlsxStylesheet2 Stylesheet { get; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;

        public XlsxWriter2(Stream stream)
        {
            _worksheets = new List<XlsxWorksheet2>();
            Stylesheet = new XlsxStylesheet2();

            _zipWriter = (ZipWriter)WriterFactory.Open(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate));
        }

        public void Dispose()
        {
            _currentWorksheet?.Dispose();
            Stylesheet.Save(_zipWriter);
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
                    worksheetTags.Append($"<sheet name=\"{worksheet.Name}\" sheetId=\"{worksheet.Id}\" r:id=\"RidWS{worksheet.Id}\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>");
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

        public XlsxWriter2 BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0)
        {
            _currentWorksheet?.Dispose();
            _currentWorksheet = new XlsxWorksheet2(_zipWriter, _worksheets.Count + 1, name, splitRow, splitColumn);
            _worksheets.Add(_currentWorksheet);
            return this;
        }

        public XlsxWriter2 SkipRows(int rowCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipRows(rowCount);
            return this;
        }

        public XlsxWriter2 BeginRow()
        {
            EnsureWorksheet();
            _currentWorksheet.BeginRow();
            return this;
        }

        public XlsxWriter2 SkipColumns(int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipColumns(columnCount);
            return this;
        }

        public XlsxWriter2 Write()
        {
            return Write(XlsxStyle2.Default);
        }

        public XlsxWriter2 Write(XlsxStyle2 style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(style);
            return this;
        }

        public XlsxWriter2 Write(string value)
        {
            return Write(value, XlsxStyle2.Default);
        }

        public XlsxWriter2 Write(string value, XlsxStyle2 style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter2 Write(double value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, XlsxStyle2.Default);
            return this;
        }

        public XlsxWriter2 Write(double value, XlsxStyle2 style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter2 Write(decimal value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write((double)value, XlsxStyle2.Default);
            return this;
        }

        public XlsxWriter2 Write(decimal value, XlsxStyle2 style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write((double)value, style);
            return this;
        }

        public XlsxWriter2 Write(int value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, XlsxStyle2.Default);
            return this;
        }

        public XlsxWriter2 Write(int value, XlsxStyle2 style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter2 AddMergedCell(int rowCount, int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.AddMergedCell(_currentWorksheet.CurrentRowNumber, _currentWorksheet.CurrentColumnNumber, rowCount, columnCount);
            return this;
        }

        public XlsxWriter2 AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.AddMergedCell(fromRow, fromColumn, rowCount, columnCount);
            return this;
        }

        private void EnsureWorksheet()
        {
            if (_currentWorksheet == null)
                throw new InvalidOperationException($"{nameof(BeginWorksheet)} not called");
        }
    }
}
