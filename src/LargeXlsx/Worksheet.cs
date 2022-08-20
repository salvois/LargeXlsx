/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2022 Salvatore ISAJA. All rights reserved.

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
using System.Globalization;
using System.IO;
using System.Linq;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    internal class Worksheet : IDisposable
    {
        private const int MinSheetProtectionPasswordLength = 1;
        private const int MaxSheetProtectionPasswordLength = 255;
        private const int MaxRowNumbers = 1048576;
        private readonly Stream _stream;
        private readonly StreamWriter _streamWriter;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly List<string> _mergedCellRefs;
        private readonly Dictionary<XlsxDataValidation, List<string>> _cellRefsByDataValidation;
        private string _autoFilterRef;
        private string _autoFilterAbsoluteRef;
        private XlsxSheetProtection _sheetProtection;

        public int Id { get; }
        public string Name { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }
        internal string AutoFilterAbsoluteRef => _autoFilterAbsoluteRef;

        public Worksheet(ZipWriter zipWriter, int id, string name, int splitRow, int splitColumn, bool rightToLeft, Stylesheet stylesheet, SharedStringTable sharedStringTable, IEnumerable<XlsxColumn> columns)
        {
            Id = id;
            Name = name;
            CurrentRowNumber = 0;
            CurrentColumnNumber = 0;
            _stylesheet = stylesheet;
            _sharedStringTable = sharedStringTable;
            _mergedCellRefs = new List<string>();
            _cellRefsByDataValidation = new Dictionary<XlsxDataValidation, List<string>>();
            _stream = zipWriter.WriteToStream($"xl/worksheets/sheet{id}.xml", new ZipWriterEntryOptions());
            _streamWriter = new InvariantCultureStreamWriter(_stream);

            _streamWriter.WriteLine("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"
                                + "<sheetViews>"
                                + $"<sheetView workbookViewId=\"0\" rightToLeft=\"{(rightToLeft ? 1 : 0)}\">");
            if (splitRow > 0 || splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
            _streamWriter.WriteLine("</sheetView></sheetViews>");
            if (columns.Any())
                WriteColumns(columns);
            _streamWriter.WriteLine("<sheetData>");
        }

        public void Dispose()
        {
            CloseLastRow();
            _streamWriter.WriteLine("</sheetData>");
            WriteSheetProtection();
            WriteAutoFilter();
            WriteMergedCells();
            WriteDataValidations();
            _streamWriter.WriteLine("</worksheet>");
            _streamWriter.Dispose();
            _stream.Dispose();
        }

        public void BeginRow(double? height, bool hidden, XlsxStyle style)
        {
            CloseLastRow();
            if (CurrentRowNumber == MaxRowNumbers)
                throw new InvalidOperationException($"A worksheet can contain at most {MaxRowNumbers} rows ({CurrentRowNumber + 1} attempted)");
            CurrentRowNumber++;
            CurrentColumnNumber = 1;
            _streamWriter.Write("<row r=\"{0}\"", CurrentRowNumber);
            if (height.HasValue) _streamWriter.Write(" ht=\"{0}\" customHeight=\"1\"", height);
            if (hidden) _streamWriter.Write(" hidden=\"1\"");
            if (style != null) _streamWriter.Write(" s=\"{0}\" customFormat=\"1\"", _stylesheet.ResolveStyleId(style));
            _streamWriter.WriteLine(">");
        }

        public void SkipRows(int rowCount)
        {
            CloseLastRow();
            if (CurrentRowNumber + rowCount > MaxRowNumbers)
                throw new InvalidOperationException($"A worksheet can contain at most {MaxRowNumbers} rows ({CurrentRowNumber + rowCount} attempted)");
            CurrentRowNumber += rowCount;
        }

        public void SkipColumns(int columnCount)
        {
            EnsureRow();
            CurrentColumnNumber += columnCount;
        }

        public void Write(XlsxStyle style, int repeatCount)
        {
            EnsureRow();
            var resolveStyleId = _stylesheet.ResolveStyleId(style);
            for (var i = 0; i < repeatCount; i++)
                _streamWriter.WriteLine("<c r=\"{0}{1}\" s=\"{2}\"/>", Util.GetColumnName(CurrentColumnNumber++), CurrentRowNumber, resolveStyleId);
        }

        public void Write(string value, XlsxStyle style)
        {
            if (value == null)
            {
                Write(style, 1);
                return;
            }

            EnsureRow();
            _streamWriter.WriteLine("<c r=\"{0}{1}\" s=\"{2}\" t=\"inlineStr\"><is><t>{3}</t></is></c>",
                Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, _stylesheet.ResolveStyleId(style), Util.EscapeXmlText(value));
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            EnsureRow();
            _streamWriter.WriteLine("<c r=\"{0}{1}\" s=\"{2}\"><v>{3}</v></c>",
                Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, _stylesheet.ResolveStyleId(style), value);
            CurrentColumnNumber++;
        }

        public void WriteFormula(string formula, XlsxStyle style, IConvertible result)
        {
            EnsureRow();
            _streamWriter.WriteLine("<c r=\"{0}{1}\" s=\"{2}\" t=\"str\"><f>{3}</f>",
                Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, _stylesheet.ResolveStyleId(style), Util.EscapeXmlText(formula));
            if (result != null) _streamWriter.Write("<v>{0}</v>", Util.EscapeXmlText(result.ToString(CultureInfo.InvariantCulture)));
            _streamWriter.Write("</c>");
            CurrentColumnNumber++;
        }

        public void WriteSharedString(string value, XlsxStyle style)
        {
            EnsureRow();
            _streamWriter.WriteLine("<c r=\"{0}{1}\" s=\"{2}\" t=\"s\"><v>{3}</v></c>",
                Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, _stylesheet.ResolveStyleId(style), _sharedStringTable.ResolveStringId(value));
            CurrentColumnNumber++;
        }

        public void AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = Util.GetColumnName(fromColumn);
            var toColumnName = Util.GetColumnName(fromColumn + columnCount - 1);
            _mergedCellRefs.Add($"{fromColumnName}{fromRow}:{toColumnName}{toRow}");
        }

        public void SetAutoFilter(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = Util.GetColumnName(fromColumn);
            var toColumnName = Util.GetColumnName(fromColumn + columnCount - 1);
            _autoFilterRef = $"{fromColumnName}{fromRow}:{toColumnName}{toRow}";
            _autoFilterAbsoluteRef = $"'{Name.Replace("'", "''")}'!${fromColumnName}${fromRow}:${toColumnName}${toRow}";
        }

        public void AddDataValidation(int fromRow, int fromColumn, int rowCount, int columnCount, XlsxDataValidation dataValidation)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var cellRef = rowCount > 1 || columnCount > 1
                ? $"{Util.GetColumnName(fromColumn)}{fromRow}:{Util.GetColumnName(fromColumn + columnCount - 1)}{fromRow + rowCount - 1}"
                : $"{Util.GetColumnName(fromColumn)}{fromRow}";
            if (!_cellRefsByDataValidation.TryGetValue(dataValidation, out var cellRefs))
            {
                cellRefs = new List<string>();
                _cellRefsByDataValidation.Add(dataValidation, cellRefs);
            }
            cellRefs.Add(cellRef);
        }

        public void SetSheetProtection(XlsxSheetProtection sheetProtection)
        {
            if (sheetProtection.Password.Length < MinSheetProtectionPasswordLength || sheetProtection.Password.Length > MaxSheetProtectionPasswordLength)
                throw new ArgumentException("Invalid password length");
            _sheetProtection = sheetProtection;
        }

        private void EnsureRow()
        {
            if (CurrentColumnNumber == 0)
                throw new InvalidOperationException($"{nameof(BeginRow)} not called");
        }

        private void CloseLastRow()
        {
            if (CurrentColumnNumber > 0)
            {
                _streamWriter.WriteLine("</row>");
                CurrentColumnNumber = 0;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{Util.GetColumnName(fromColumn + 1)}{fromRow + 1}";
            if (fromRow > 0 && fromColumn > 0)
            {
                _streamWriter.WriteLine("<pane xSplit=\"{0}\" ySplit=\"{1}\" topLeftCell=\"{2}\" activePane=\"bottomRight\" state=\"frozen\"/>"
                                        + "<selection pane=\"bottomRight\" activeCell=\"{2}\" sqref=\"{2}\"/>",
                    fromColumn, fromRow, topLeftCell);
            }
            else if (fromRow > 0)
            {
                _streamWriter.WriteLine("<pane ySplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"bottomLeft\" state=\"frozen\"/>"
                                    + "<selection pane=\"bottomLeft\" activeCell=\"{1}\" sqref=\"{1}\"/>",
                    fromRow, topLeftCell);
            }
            else if (fromColumn > 0)
            {
                _streamWriter.WriteLine("<pane xSplit=\"{0}\" topLeftCell=\"{1}\" activePane=\"topRight\" state=\"frozen\"/>"
                                    + "<selection pane=\"topRight\" activeCell=\"{1}\" sqref=\"{1}\"/>",
                    fromColumn, topLeftCell);
            }
        }

        private void WriteColumns(IEnumerable<XlsxColumn> columns)
        {
            var columnIndex = 1;
            _streamWriter.Write("<cols>");
            foreach (var column in columns)
            {
                if (column.Hidden || column.Style != null || column.Width.HasValue)
                {
                    _streamWriter.Write("<col min=\"{0}\" max=\"{1}\"", columnIndex, columnIndex + column.Count - 1);
                    if (column.Width.HasValue) _streamWriter.Write(" width=\"{0}\"", column.Width.Value);
                    if (column.Hidden) _streamWriter.Write(" hidden=\"1\"");
                    if (column.Width.HasValue) _streamWriter.Write(" customWidth=\"1\"");
                    if (column.Style != null) _streamWriter.Write(" style=\"{0}\"", _stylesheet.ResolveStyleId(column.Style));
                    _streamWriter.WriteLine("/>");
                }
                columnIndex += column.Count;
            }
            _streamWriter.WriteLine("</cols>");
        }

        private void WriteAutoFilter()
        {
            if (_autoFilterRef != null)
                _streamWriter.WriteLine("<autoFilter ref=\"{0}\"/>", _autoFilterRef);
        }

        private void WriteMergedCells()
        {
            if (!_mergedCellRefs.Any())
                return;
            _streamWriter.WriteLine("<mergeCells count=\"{0}\">", _mergedCellRefs.Count);
            foreach (var mergedCell in _mergedCellRefs)
                _streamWriter.WriteLine("<mergeCell ref=\"{0}\"/>", mergedCell);
            _streamWriter.WriteLine("</mergeCells>");
        }

        private void WriteDataValidations()
        {
            if (!_cellRefsByDataValidation.Any())
                return;
            _streamWriter.WriteLine("<dataValidations count=\"{0}\">", _cellRefsByDataValidation.Count);
            foreach (var kvp in _cellRefsByDataValidation)
            {
                _streamWriter.Write("<dataValidation sqref=\"{0}\" allowBlank=\"{1}\"",
                    string.Join(" ", kvp.Value.Distinct()), Util.BoolToInt(kvp.Key.AllowBlank));
                if (kvp.Key.Error != null) _streamWriter.Write(" error=\"{0}\"", Util.EscapeXmlAttribute(kvp.Key.Error));
                if (kvp.Key.ErrorStyleValue.HasValue) _streamWriter.Write(" errorStyle=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.ErrorStyleValue));
                if (kvp.Key.ErrorTitle != null) _streamWriter.Write(" errorTitle=\"{0}\"", Util.EscapeXmlAttribute(kvp.Key.ErrorTitle));
                if (kvp.Key.OperatorValue.HasValue) _streamWriter.Write(" operator=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.OperatorValue));
                if (kvp.Key.Prompt != null) _streamWriter.Write(" prompt=\"{0}\"", Util.EscapeXmlAttribute(kvp.Key.Prompt));
                if (kvp.Key.PromptTitle != null) _streamWriter.Write(" promptTitle=\"{0}\"", Util.EscapeXmlAttribute(kvp.Key.PromptTitle));
                if (kvp.Key.ShowDropDown) _streamWriter.Write(" showDropDown=\"1\"");
                if (kvp.Key.ShowErrorMessage) _streamWriter.Write(" showErrorMessage=\"1\"");
                if (kvp.Key.ShowInputMessage) _streamWriter.Write(" showInputMessage=\"1\"");
                if (kvp.Key.ValidationTypeValue.HasValue) _streamWriter.Write(" type=\"{0}\"", Util.EnumToAttributeValue(kvp.Key.ValidationTypeValue));
                _streamWriter.Write(">");
                if (kvp.Key.Formula1 != null) _streamWriter.Write("<formula1>{0}</formula1>", Util.EscapeXmlText(kvp.Key.Formula1));
                if (kvp.Key.Formula2 != null) _streamWriter.Write("<formula2>{0}</formula2>", Util.EscapeXmlText(kvp.Key.Formula2));
                _streamWriter.WriteLine("</dataValidation>");
            }
            _streamWriter.WriteLine("</dataValidations>");
        }

        private void WriteSheetProtection()
        {
            if (_sheetProtection == null)
                return;
            const int spinCount = 100000;
            var saltValue = Guid.NewGuid().ToByteArray();
            var hash = Util.ComputePasswordHash(_sheetProtection.Password, saltValue, spinCount);
            _streamWriter.Write("<sheetProtection algorithmName=\"SHA-512\" hashValue=\"{0}\" saltValue=\"{1}\" spinCount=\"{2}\"", Convert.ToBase64String(hash), Convert.ToBase64String(saltValue), spinCount);
            if (_sheetProtection.Sheet) _streamWriter.Write(" sheet=\"1\"");
            if (_sheetProtection.Objects) _streamWriter.Write(" objects=\"1\"");
            if (_sheetProtection.Scenarios) _streamWriter.Write(" scenarios=\"1\"");
            if (!_sheetProtection.FormatCells) _streamWriter.Write(" formatCells=\"0\"");
            if (!_sheetProtection.FormatColumns) _streamWriter.Write(" formatColumns=\"0\"");
            if (!_sheetProtection.FormatRows) _streamWriter.Write(" formatRows=\"0\"");
            if (!_sheetProtection.InsertColumns) _streamWriter.Write(" insertColumns=\"0\"");
            if (!_sheetProtection.InsertRows) _streamWriter.Write(" insertRows=\"0\"");
            if (!_sheetProtection.InsertHyperlinks) _streamWriter.Write(" insertHyperlinks=\"0\"");
            if (!_sheetProtection.DeleteColumns) _streamWriter.Write(" deleteColumns=\"0\"");
            if (!_sheetProtection.DeleteRows) _streamWriter.Write(" deleteRows=\"0\"");
            if (_sheetProtection.SelectLockedCells) _streamWriter.Write(" selectLockedCells=\"1\"");
            if (!_sheetProtection.Sort) _streamWriter.Write(" sort=\"0\"");
            if (!_sheetProtection.AutoFilter) _streamWriter.Write(" autoFilter=\"0\"");
            if (!_sheetProtection.PivotTables) _streamWriter.Write(" pivotTables=\"0\"");
            if (_sheetProtection.SelectUnlockedCells) _streamWriter.Write(" selectUnlockedCells=\"1\"");
            _streamWriter.WriteLine("/>");
        }
    }
}