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
        private readonly List<XlsxSheet2> _largeXlsxSheets;
        private XlsxSheet2 _currentSheet;

        public XlsxStylesheet Stylesheet { get; }

        public XlsxWriter2(Stream stream)
        {
            _largeXlsxSheets = new List<XlsxSheet2>();
            Stylesheet = new XlsxStylesheet();

            _zipWriter = (ZipWriter)WriterFactory.Open(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate));
        }

        public void Dispose()
        {
            _currentSheet?.Dispose();

            //Stylesheet.Save(_document);

            using (var stream = _zipWriter.WriteToStream("[Content_Types].xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                                   + "<Default Extension=\"xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />"
                                   + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />"
                                   + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />"
                                   + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />"
                                   + "</Types>");
            }
            using (var stream = _zipWriter.WriteToStream("_rels/.rels", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"/xl/workbook.xml\" Id=\"R7130d2a4b4db42e0\" />"
                                   + "</Relationships>");
            }
            using (var stream = _zipWriter.WriteToStream("xl/workbook.xml", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                var sheetTags = new StringBuilder();
                var sheetId = 1;
                // TODO: r:id is hardcoded
                foreach (var sheet in _largeXlsxSheets)
                    sheetTags.Append($"<sheet name=\"{sheet.Name}\" sheetId=\"{sheetId++}\" r:id=\"Rc3797908a4cd4249\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                   + "<sheets>"
                                   + sheetTags
                                   + "</sheets>"
                                   + "</workbook>");
            }
            using (var stream = _zipWriter.WriteToStream("xl/_rels/workbook.xml.rels", new ZipWriterEntryOptions()))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet1.xml\" Id=\"Rc3797908a4cd4249\" />"
                                   //+ "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"Rb18eccf29de245a8\" />"
                                   + "</Relationships>");
            }
            _zipWriter.Dispose();
        }

        public XlsxWriter2 BeginSheet(string name, int splitRow = 0, int splitColumn = 0)
        {
            _currentSheet?.Dispose();
            _currentSheet = new XlsxSheet2(_zipWriter, name, splitRow, splitColumn);
            _largeXlsxSheets.Add(_currentSheet);
            return this;
        }

        public XlsxWriter2 SkipRows(int rowCount)
        {
            EnsureSheet();
            _currentSheet.SkipRows(rowCount);
            return this;
        }

        public XlsxWriter2 BeginRow()
        {
            EnsureSheet();
            _currentSheet.BeginRow();
            return this;
        }

        public XlsxWriter2 SkipColumns(int columnCount)
        {
            EnsureSheet();
            _currentSheet.SkipColumns(columnCount);
            return this;
        }

        public XlsxWriter2 WriteInlineStringCell(string value)
        {
            return WriteInlineStringCell(value, XlsxStylesheet2.DefaultStyle);
        }

        public XlsxWriter2 WriteInlineStringCell(string value, XlsxStyle style)
        {
            EnsureSheet();
            _currentSheet.WriteInlineStringCell(value, style);
            return this;
        }

        public XlsxWriter2 WriteNumericCell(double value)
        {
            EnsureSheet();
            _currentSheet.WriteNumericCell(value, XlsxStylesheet2.DefaultStyle);
            return this;
        }

        public XlsxWriter2 WriteNumericCell(double value, XlsxStyle style)
        {
            EnsureSheet();
            _currentSheet.WriteNumericCell(value, style);
            return this;
        }

        public XlsxWriter2 AddMergedCell(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            EnsureSheet();
            _currentSheet.AddMergedCell(fromRow, fromColumn, toRow, toColumn);
            return this;
        }

        private void EnsureSheet()
        {
            if (_currentSheet == null)
                throw new InvalidOperationException($"{nameof(BeginSheet)} not called");
        }
    }
}