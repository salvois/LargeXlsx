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
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Text;

namespace LargeXlsx
{
    public sealed class XlsxWriter : IDisposable
    {
        private const int MaxSheetNameLength = 31;
        private readonly ZipArchive _zipArchive;
        private readonly List<Worksheet> _worksheets;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly bool _requireCellReferences;
        private readonly bool _skipInvalidCharacters;
        private Worksheet _currentWorksheet;
        private bool _hasFormulasWithoutResult;
        private bool _disposed;

        public XlsxStyle DefaultStyle { get; private set; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;
        public string CurrentColumnName => Util.GetColumnName(CurrentColumnNumber);
        public string GetRelativeColumnName(int offsetFromCurrentColumn) => Util.GetColumnName(CurrentColumnNumber + offsetFromCurrentColumn);
        public static string GetColumnName(int columnIndex) => Util.GetColumnName(columnIndex);

        public XlsxWriter(Stream stream, CompressionLevel compressionLevel = CompressionLevel.Optimal, bool useZip64 = false, bool requireCellReferences = true, bool skipInvalidCharacters = false)
        {
            _worksheets = new List<Worksheet>();
            _stylesheet = new Stylesheet();
            _sharedStringTable = new SharedStringTable(skipInvalidCharacters);
            _requireCellReferences = requireCellReferences;
            _skipInvalidCharacters = skipInvalidCharacters;
            DefaultStyle = XlsxStyle.Default;

            _zipArchive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true);
        }

        public void Dispose()
        {
            if (!_disposed)
            {
                _currentWorksheet?.Dispose();
                _stylesheet.Save(_zipArchive);
                _sharedStringTable.Save(_zipArchive);
                SaveDocProps();
                SaveContentTypes();
                SaveRels();
                SaveWorkbook();
                SaveWorkbookRels();
                _zipArchive.Dispose();
                _disposed = true;
            }
        }

        private void SaveDocProps()
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName();
            var entry = _zipArchive.CreateEntry("docProps/app.xml", CompressionLevel.Optimal);
            using (var streamWriter = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                // Looks some applications (e.g. Microsoft's) may consider a file invalid if a specific version number is not found.
                // Thus, pretend being version 15.0 like LibreOffice Calc does.
                // https://bugs.documentfoundation.org/show_bug.cgi?id=91064
                streamWriter
                    .Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                            + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
                            + "<Application>").AppendEscapedXmlText(assemblyName.Name, false)
                    .Append($"/{assemblyName.Version.Major}.{assemblyName.Version.Minor}.{assemblyName.Version.Build}</Application>"
                            + "<AppVersion>15.0000</AppVersion>"
                            + "</Properties>");
            }
        }

        private void SaveContentTypes()
        {
            var entry = _zipArchive.CreateEntry("[Content_Types].xml", CompressionLevel.Optimal);
            using (var streamWriter = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                var worksheetTags = new StringBuilder();
                foreach (var worksheet in _worksheets)
                    worksheetTags.Append($"<Override PartName=\"/xl/worksheets/sheet{worksheet.Id}.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>");
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                                   + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                                   + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "<Override PartName=\"/_rels/.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>"
                                   + "<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>"
                                   + worksheetTags
                                   + "<Override PartName=\"/xl/sharedStrings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>"
                                   + "<Override PartName=\"/xl/_rels/workbook.xml.rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                                   + "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>"
                                   + "</Types>");
            }
        }

        private void SaveRels()
        {
            var entry = _zipArchive.CreateEntry("_rels/.rels", CompressionLevel.Optimal);
            using (var streamWriter = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Id=\"rIdWb1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                                   + "<Relationship Id=\"app\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
                                   + "</Relationships>");
            }
        }

        private void SaveWorkbook()
        {
            var entry = _zipArchive.CreateEntry("xl/workbook.xml", CompressionLevel.Optimal);
            using (var streamWriter = new StreamWriter(entry.Open(), Encoding.UTF8))
            {
                var worksheetTags = new StringWriter();
                var definedNames = new StringWriter();
                var sheetIndex = 0;
                foreach (var worksheet in _worksheets)
                {
                    worksheetTags
                        .Append("<sheet name=\"")
                        .AppendEscapedXmlAttribute(worksheet.Name, _skipInvalidCharacters)
                        .Append($"\" sheetId=\"{worksheet.Id}\" {GetWorksheetState(worksheet.State)} r:id=\"RidWS{worksheet.Id}\"/>");
                    if (worksheet.AutoFilterAbsoluteRef != null)
                        definedNames
                            .Append($"<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\"{sheetIndex}\" hidden=\"1\">")
                            .AppendEscapedXmlText(worksheet.AutoFilterAbsoluteRef, _skipInvalidCharacters)
                            .Append("</definedName>");
                    sheetIndex++;
                }
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
                                   + "<sheets>"
                                   + worksheetTags
                                   + "</sheets>");
                var definedNamesString = definedNames.ToString();
                if (definedNamesString.Length > 0)
                    streamWriter.Append("<definedNames>").Append(definedNamesString).Append("</definedNames>");
                if (_hasFormulasWithoutResult) streamWriter.Write("<calcPr calcCompleted=\"0\" fullCalcOnLoad=\"1\"/>");
                streamWriter.Write("</workbook>");
            }
        }

        private static string GetWorksheetState(XlsxWorksheetState state)
        {
            switch (state)
            {
                case XlsxWorksheetState.Visible:
                    return "";
                case XlsxWorksheetState.Hidden:
                    return "state=\"hidden\"";
                case XlsxWorksheetState.VeryHidden:
                    return "state=\"veryHidden\"";
                default:
                    throw new ArgumentOutOfRangeException(nameof(state), state, null);
            }
        }

        private void SaveWorkbookRels()
        {
            var entry = _zipArchive.CreateEntry("xl/_rels/workbook.xml.rels", CompressionLevel.Optimal);
            using (var streamWriter = new StreamWriter(entry.Open(), Encoding.UTF8))
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

        public XlsxWriter BeginWorksheet(
            string name,
            int splitRow = 0,
            int splitColumn = 0,
            bool rightToLeft = false,
            IEnumerable<XlsxColumn> columns = null,
            bool showGridLines = true,
            bool showHeaders = true,
            XlsxWorksheetState state = XlsxWorksheetState.Visible)
        {
            if (name.Length > MaxSheetNameLength)
                throw new ArgumentException($"The name \"{name}\" exceeds the maximum length of {MaxSheetNameLength} characters supported by Excel");
            if (_worksheets.Any(ws => string.Equals(ws.Name, name, StringComparison.InvariantCultureIgnoreCase)))
                throw new ArgumentException($"A worksheet named \"{name}\" has already been added");
            _currentWorksheet?.Dispose();
            _currentWorksheet = new Worksheet(
                zipArchive: _zipArchive,
                id: _worksheets.Count + 1,
                name: name,
                splitRow: splitRow,
                splitColumn: splitColumn,
                rightToLeft: rightToLeft,
                state: state,
                stylesheet: _stylesheet,
                sharedStringTable: _sharedStringTable,
                columns: columns ?? Enumerable.Empty<XlsxColumn>(),
                showGridLines: showGridLines,
                showHeaders: showHeaders,
                requireCellReferences: _requireCellReferences,
                skipInvalidCharacters: _skipInvalidCharacters);
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
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
        }

        public XlsxWriter Write(int value, XlsxStyle style = null, int columnSpan = 1)
        {
            if (columnSpan == 1)
            {
                CheckInWorksheet();
                _currentWorksheet.Write(value, style ?? DefaultStyle);
                return this;
            }

            return AddMergedCell(1, columnSpan).Write(value, style, 1).Write(style, repeatCount: columnSpan - 1);
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

        public XlsxWriter AddMergedCell(int rowCount, int columnCount) =>
            AddMergedCell(CurrentRowNumber, CurrentColumnNumber, rowCount, columnCount);

        public XlsxWriter AddRowPageBreakBefore(int rowNumber)
        {
            CheckInWorksheet();
            _currentWorksheet.AddRowPageBreakBefore(rowNumber);
            return this;
        }

        public XlsxWriter AddColumnPageBreakBefore(int columnNumber)
        {
            CheckInWorksheet();
            _currentWorksheet.AddColumnPageBreakBefore(columnNumber);
            return this;
        }

        public XlsxWriter AddRowPageBreak() =>
            AddRowPageBreakBefore(CurrentRowNumber);

        public XlsxWriter AddColumnPageBreak() =>
            AddColumnPageBreakBefore(CurrentColumnNumber);

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

        public XlsxWriter SetHeaderFooter(XlsxHeaderFooter headerFooter)
        {
            CheckInWorksheet();
            _currentWorksheet.SetHeaderFooter(headerFooter);
            return this;
        }

        private void CheckInWorksheet()
        {
            if (_currentWorksheet == null)
                throw new InvalidOperationException($"{nameof(BeginWorksheet)} not called");
        }
    }
}
