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
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace LargeXlsx
{
    public sealed class XlsxWriter : IDisposable, IAsyncDisposable
    {
        private const int MaxSheetNameLength = 31;
        private readonly IZipWriter _zipWriter;
        private readonly List<Worksheet> _worksheets;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly bool _requireCellReferences;
        private readonly bool _skipInvalidCharacters;
        private readonly CustomWriter _customWriter;
        private Worksheet _currentWorksheet;
        private bool _hasFormulasWithoutResult;
        private bool _disposed;

        public XlsxStyle DefaultStyle { get; private set; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;
        public string CurrentColumnName => Util.GetColumnName(CurrentColumnNumber);
        public int BufferCapacity => _customWriter.WriteBufferCapacity;

        public string GetRelativeColumnName(int offsetFromCurrentColumn) =>
            Util.GetColumnName(CurrentColumnNumber + offsetFromCurrentColumn);

        public static string GetColumnName(int columnIndex) =>
            Util.GetColumnName(columnIndex);

        public XlsxWriter(Stream stream, XlsxCompressionLevel compressionLevel = XlsxCompressionLevel.Fastest, bool requireCellReferences = true, bool skipInvalidCharacters = false, int commitThreshold = 57344)
            : this(
#if NETCOREAPP2_1_OR_GREATER
                new SystemIoCompressionZipWriter(stream, compressionLevel),
#else
                new SharpCompressZipWriter(stream, compressionLevel, useZip64: true),
#endif
                requireCellReferences: requireCellReferences,
                skipInvalidCharacters: skipInvalidCharacters,
                commitThreshold: commitThreshold)
        {
        }

        internal XlsxWriter(IZipWriter zipWriter, bool requireCellReferences = true, bool skipInvalidCharacters = false, int commitThreshold = 65536)
        {
            _worksheets = new List<Worksheet>();
            _stylesheet = new Stylesheet();
            _sharedStringTable = new SharedStringTable(skipInvalidCharacters);
            _requireCellReferences = requireCellReferences;
            _skipInvalidCharacters = skipInvalidCharacters;
            _customWriter = new CustomWriter(commitThreshold);
            DefaultStyle = XlsxStyle.Default;
            _zipWriter = zipWriter;
        }

        public XlsxWriter TryCommit()
        {
            CheckInWorksheet();
            _currentWorksheet.TryCommit();
            return this;
        }

        public XlsxWriter Commit()
        {
            CheckInWorksheet();
            _currentWorksheet.Commit();
            return this;
        }

        public async Task<XlsxWriter> TryCommitAsync()
        {
            CheckInWorksheet();
            await _currentWorksheet.TryCommitAsync();
            return this;
        }

        public async Task<XlsxWriter> CommitAsync()
        {
            CheckInWorksheet();
            await _currentWorksheet.CommitAsync();
            return this;
        }

        public void Dispose() => 
            DisposeAsync().GetAwaiter().GetResult();

        public async ValueTask DisposeAsync()
        {
            if (!_disposed)
            {
                if (_currentWorksheet != null)
                    await _currentWorksheet.DisposeAsync().ConfigureAwait(false);
                await _stylesheet.Save(_zipWriter, _customWriter).ConfigureAwait(false);
                await _sharedStringTable.Save(_zipWriter, _customWriter).ConfigureAwait(false);
                SaveDocProps();
                SaveContentTypes();
                SaveRels();
                await SaveWorkbook().ConfigureAwait(false);
                SaveWorkbookRels();
                _zipWriter.Dispose();
                _disposed = true;
            }
        }

        private void SaveDocProps()
        {
            var assemblyName = Assembly.GetExecutingAssembly().GetName();
            using (var stream = _zipWriter.CreateEntry("docProps/app.xml"))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                // Looks some applications (e.g. Microsoft's) may consider a file invalid if a specific version number is not found.
                // Thus, pretend being version 15.0 like LibreOffice Calc does.
                // https://bugs.documentfoundation.org/show_bug.cgi?id=91064
                streamWriter
                    .Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                            + "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">"
                            + $"<Application>{assemblyName.Name}/{assemblyName.Version.Major}.{assemblyName.Version.Minor}.{assemblyName.Version.Build}</Application>"
                            + "<AppVersion>15.0000</AppVersion>"
                            + "</Properties>");
            }
        }

        private void SaveContentTypes()
        {
            using (var stream = _zipWriter.CreateEntry("[Content_Types].xml"))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
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
            using (var stream = _zipWriter.CreateEntry("_rels/.rels"))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
            {
                streamWriter.Write("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"
                                   + "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">"
                                   + "<Relationship Id=\"rIdWb1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/>"
                                   + "<Relationship Id=\"app\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" Target=\"docProps/app.xml\"/>"
                                   + "</Relationships>");
            }
        }

        private async Task SaveWorkbook()
        {
#if NETCOREAPP2_1_OR_GREATER
            await using var stream = _zipWriter.CreateEntry("xl/workbook.xml");
#else
            using var stream = _zipWriter.CreateEntry("xl/workbook.xml");
#endif
            _customWriter.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8
                                 + "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"u8
                                 + "<sheets>"u8);
            var hasDefinedNames = false;
            for (var i = 0; i < _worksheets.Count; i++)
            {
                var worksheet = _worksheets[i];
                _customWriter
                    .Append("<sheet name=\""u8)
                    .AppendEscapedXmlAttribute(worksheet.Name, _skipInvalidCharacters)
                    .Append("\" sheetId=\""u8)
                    .Append(worksheet.Id)
                    .Append("\" "u8)
                    .Append(GetWorksheetState(worksheet.State))
                    .Append(" r:id=\"RidWS"u8)
                    .Append(worksheet.Id)
                    .Append("\"/>"u8);
                if (worksheet.AutoFilterAbsoluteRef != null)
                    hasDefinedNames = true;
            }
            _customWriter.Append("</sheets>"u8);
            if (hasDefinedNames)
            {
                _customWriter.Append("<definedNames>"u8);
                for (var i = 0; i < _worksheets.Count; i++)
                {
                    var worksheet = _worksheets[i];
                    if (worksheet.AutoFilterAbsoluteRef != null)
                        _customWriter
                            .Append("<definedName name=\"_xlnm._FilterDatabase\" localSheetId=\""u8)
                            .Append(i)
                            .Append("\" hidden=\"1\">"u8)
                            .AppendEscapedXmlText(worksheet.AutoFilterAbsoluteRef, _skipInvalidCharacters)
                            .Append("</definedName>"u8);
                }
                _customWriter.Append("</definedNames>"u8);
            }
            if (_hasFormulasWithoutResult)
                _customWriter.Append("<calcPr calcCompleted=\"0\" fullCalcOnLoad=\"1\"/>"u8);
            await _customWriter.Append("</workbook>"u8).FlushToAsync(stream).ConfigureAwait(false);
        }

        private static ReadOnlySpan<byte> GetWorksheetState(XlsxWorksheetState state)
        {
            switch (state)
            {
                case XlsxWorksheetState.Visible:
                    return ""u8;
                case XlsxWorksheetState.Hidden:
                    return "state=\"hidden\""u8;
                case XlsxWorksheetState.VeryHidden:
                    return "state=\"veryHidden\""u8;
                default:
                    throw new ArgumentOutOfRangeException(nameof(state), state, null);
            }
        }

        private void SaveWorkbookRels()
        {
            using (var stream = _zipWriter.CreateEntry("xl/_rels/workbook.xml.rels"))
            using (var streamWriter = new StreamWriter(stream, Encoding.UTF8))
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
            return BeginWorksheetAsync(name, splitRow, splitColumn, rightToLeft, columns, showGridLines, showHeaders, state).GetAwaiter().GetResult();
        }
        
        public async Task<XlsxWriter> BeginWorksheetAsync(
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
            if (_currentWorksheet != null)
                await _currentWorksheet.DisposeAsync().ConfigureAwait(false);
            _currentWorksheet = new Worksheet(
                zipWriter: _zipWriter,
                customWriter: _customWriter,
                id: _worksheets.Count + 1,
                name: name,
                splitRow: splitRow,
                splitColumn: splitColumn,
                rightToLeft: rightToLeft,
                state: state,
                stylesheet: _stylesheet,
                sharedStringTable: _sharedStringTable,
                columns: columns ?? [],
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
            TryCommit();
            _currentWorksheet.BeginRow(height, hidden, style);
            return this;
        }

        public async Task<XlsxWriter> BeginRowAsync(double? height = null, bool hidden = false, XlsxStyle style = null)
        {
            CheckInWorksheet();
            await TryCommitAsync().ConfigureAwait(false);
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
