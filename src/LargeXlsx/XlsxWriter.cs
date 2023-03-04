/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2023 Salvatore ISAJA. All rights reserved.

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
using ICSharpCode.SharpZipLib.Zip;

namespace LargeXlsx
{
    public sealed class XlsxWriter : IDisposable
    {
        private const int MaxSheetNameLength = 31;
        private readonly ZipOutputStream _zipOutputStream;
        private readonly List<Worksheet> _worksheets;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private Worksheet _currentWorksheet;
        private bool _hasFormulasWithoutResult;
        private bool _disposed;

        public XlsxStyle DefaultStyle { get; private set; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;
        public string CurrentColumnName => Util.GetColumnName(CurrentColumnNumber);
        public string GetRelativeColumnName(int offsetFromCurrentColumn) => Util.GetColumnName(CurrentColumnNumber + offsetFromCurrentColumn);
        public static string GetColumnName(int columnIndex) => Util.GetColumnName(columnIndex);

        public XlsxWriter(Stream stream, XlsxCompressionLevel compressionLevel = XlsxCompressionLevel.Level2, bool useZip64 = false)
        {
            _worksheets = new List<Worksheet>();
            _stylesheet = new Stylesheet();
            _sharedStringTable = new SharedStringTable();
            DefaultStyle = XlsxStyle.Default;
            _zipOutputStream = new ZipOutputStream(stream);
            _zipOutputStream.UseZip64 = useZip64 ? UseZip64.Dynamic : UseZip64.Off;
            _zipOutputStream.IsStreamOwner = false;
            _zipOutputStream.SetLevel((int)compressionLevel);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _currentWorksheet?.Dispose();
                _stylesheet.Save(_zipOutputStream);
                _sharedStringTable.Save(_zipOutputStream);
                Save();
                _zipOutputStream.Dispose();
                _disposed = true;
            }
        }

        private void Save()
        {
            _zipOutputStream.PutNextEntry(new ZipEntry("[Content_Types].xml"));
            using (var streamWriter = new InvariantCultureStreamWriter(_zipOutputStream))
            {
                var worksheetTags = new StringBuilder();
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<Override PartName=\"/xl/worksheets/sheet{worksheet.Id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                                   + "<Default Extension=\"xml\" ContentType=\"xml\"/>"
                                   + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                                   + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                                   + worksheetTags
                                   + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                                   + "<Override PartName=\"/xl/_rels/workbook.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "</Types>");
            }

            _zipOutputStream.PutNextEntry(new ZipEntry("_rels/.rels"));
            using (var streamWriter = new InvariantCultureStreamWriter(_zipOutputStream))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Id=\"rIdWb1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                                   + "</Relationships>");
            }

            _zipOutputStream.PutNextEntry(new ZipEntry("xl/workbook.xml"));
            using (var streamWriter = new InvariantCultureStreamWriter(_zipOutputStream))
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
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                                   + "<sheets>"
                                   + worksheetTags
                                   + "</sheets>");
                if (definedNames.Length > 0) streamWriter.Write("<definedNames>" + definedNames + "</definedNames>");
                if (_hasFormulasWithoutResult) streamWriter.Write("<calcPr calcCompleted=\"0\" fullCalcOnLoad=\"1\"/>");
                streamWriter.Write("</workbook>");
            }

            _zipOutputStream.PutNextEntry(new ZipEntry("xl/_rels/workbook.xml.rels"));
            using (var streamWriter = new InvariantCultureStreamWriter(_zipOutputStream))
            {
                var worksheetTags = new StringBuilder();
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"/xl/worksheets/sheet{worksheet.Id}.xml\" Id=\"RidWS{worksheet.Id}\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + worksheetTags
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"/xl/styles.xml\" Id=\"RidSt1\"/>"
                                   + "<Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" Target=\"/xl/sharedStrings.xml\" Id=\"RidShS1\"/>"
                                   + "</Relationships>");
            }
        }

        public XlsxWriter BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0, bool rightToLeft = false, IReadOnlyCollection<XlsxColumn> columns = null)
        {
            if (name.Length > MaxSheetNameLength)
                throw new ArgumentException($"The name \"{name}\" exceeds the maximum length of {MaxSheetNameLength} characters supported by Excel");
            if (_worksheets.Any(ws => string.Equals(ws.Name, name, StringComparison.InvariantCultureIgnoreCase)))
                throw new ArgumentException($"A worksheet named \"{name}\" has already been added");
            _currentWorksheet?.Dispose();
            _currentWorksheet = new Worksheet(_zipOutputStream, _worksheets.Count + 1, name, splitRow, splitColumn, rightToLeft, _stylesheet, _sharedStringTable, columns ?? Array.Empty<XlsxColumn>());
            _worksheets.Add(_currentWorksheet);
            return this;
        }

        public XlsxWriter SkipRows(int rowCount)
        {
            CheckInWorksheet();
            _currentWorksheet.SkipRows(rowCount);
            return this;
        }

        public XlsxWriter BeginRow(double? height = null, bool hidden = false, XlsxStyle style = null)
        {
            CheckInWorksheet();
            _currentWorksheet.BeginRow(height, hidden, style);
            return this;
        }

        public XlsxWriter SkipColumns(int columnCount)
        {
            CheckInWorksheet();
            _currentWorksheet.SkipColumns(columnCount);
            return this;
        }

        public XlsxWriter Write(XlsxStyle style = null, int columnSpan = 1, int repeatCount = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(style ?? DefaultStyle, repeatCount);
                return this;
            }

            for (var i = 0; i < repeatCount; i++)
                AddMergedCell(1, columnSpan).Write(style, 1).Write(style, repeatCount: columnSpan - 1);
            return this;
        }

        public XlsxWriter Write(string value, XlsxStyle style = null, int columnSpan = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter Write(double value, XlsxStyle style = null, int columnSpan = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
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

        public XlsxWriter Write(bool value, XlsxStyle style = null, int columnSpan = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter WriteFormula(string formula, XlsxStyle style = null, int columnSpan = 1, IConvertible result = null)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                if (result == null) _hasFormulasWithoutResult = true;
                _currentWorksheet.WriteFormula(formula, style ?? DefaultStyle, result);
                return this;
            }

            return AddMergedCell(1, columnSpan).WriteFormula(formula, style, 1, result).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter WriteSharedString(string value, XlsxStyle style = null, int columnSpan = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.WriteSharedString(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).WriteSharedString(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            CheckInWorksheet();
            _currentWorksheet.AddMergedCell(fromRow, fromColumn, rowCount, columnCount);
            return this;
        }

        public XlsxWriter AddMergedCell(int rowCount, int columnCount)
        {
            return AddMergedCell(CurrentRowNumber, CurrentColumnNumber, rowCount, columnCount);
        }

        public XlsxWriter SetAutoFilter(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            CheckInWorksheet();
            _currentWorksheet.SetAutoFilter(fromRow, fromColumn, rowCount, columnCount);
            return this;
        }

        public XlsxWriter AddDataValidation(int fromRow, int fromColumn, int rowCount, int columnCount, XlsxDataValidation dataValidation)
        {
            CheckInWorksheet();
            _currentWorksheet.AddDataValidation(fromRow, fromColumn, rowCount, columnCount, dataValidation);
            return this;
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

        public XlsxWriter SetSheetProtection(XlsxSheetProtection sheetProtection)
        {
            CheckInWorksheet();
            _currentWorksheet.SetSheetProtection(sheetProtection);
            return this;
        }

        private void CheckInWorksheet()
        {
            if (_currentWorksheet == null)
                throw new InvalidOperationException($"{nameof(BeginWorksheet)} not called");
        }
    }
}
