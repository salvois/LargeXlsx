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
using System.Globalization;
using System.IO;
using System.Linq;

namespace LargeXlsx
{
    internal class Worksheet : IDisposable
    {
        private readonly Stream _stream;
        private readonly CustomWriter _customWriter;
        private readonly Stylesheet _stylesheet;
        private readonly SharedStringTable _sharedStringTable;
        private readonly bool _requireCellReferences;
        private readonly bool _skipInvalidCharacters;
        private readonly List<string> _mergedCellRefs;
        private readonly Dictionary<XlsxDataValidation, List<string>> _cellRefsByDataValidation;
        private readonly HashSet<int> _pageBreakRowNumbers;
        private readonly HashSet<int> _pageBreakColumnNumbers;
        private string _autoFilterRef;
        private string _autoFilterAbsoluteRef;
        private XlsxSheetProtection _sheetProtection;
        private XlsxHeaderFooter _headerFooter;
        private bool _needsRef;
        private readonly byte[] _stringedCurrentRowNumber;
        private int _stringedCurrentRowNumberLength;

        public int Id { get; }
        public string Name { get; }
        public XlsxWorksheetState State { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }
        internal string AutoFilterAbsoluteRef => _autoFilterAbsoluteRef;

        public Worksheet(
            IZipWriter zipWriter,
            CustomWriter customWriter,
            int id,
            string name,
            int splitRow,
            int splitColumn,
            bool rightToLeft,
            Stylesheet stylesheet,
            SharedStringTable sharedStringTable,
            IEnumerable<XlsxColumn> columns,
            bool showGridLines,
            bool showHeaders,
            bool requireCellReferences,
            bool skipInvalidCharacters,
            XlsxWorksheetState state)
        {
            Id = id;
            Name = name;
            State = state;
            CurrentRowNumber = 0;
            CurrentColumnNumber = 0;
            _stylesheet = stylesheet;
            _sharedStringTable = sharedStringTable;
            _requireCellReferences = requireCellReferences;
            _skipInvalidCharacters = skipInvalidCharacters;
            _mergedCellRefs = new List<string>();
            _pageBreakRowNumbers = new HashSet<int>();
            _pageBreakColumnNumbers = new HashSet<int>();
            _cellRefsByDataValidation = new Dictionary<XlsxDataValidation, List<string>>();
            _stringedCurrentRowNumber = new byte[10];
            _stringedCurrentRowNumberLength = 0;
            _stream = zipWriter.CreateEntry($"xl/worksheets/sheet{id}.xml");
            _customWriter = customWriter;

            _customWriter
                .Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>"u8
                        + "<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">"u8
                        + "<sheetViews>"u8
                        + "<sheetView showGridLines=\""u8)
                .Append(Util.BoolToInt(showGridLines))
                .Append("\" showRowColHeaders=\""u8)
                .Append(Util.BoolToInt(showHeaders))
                .Append("\" rightToLeft=\""u8)
                .Append(Util.BoolToInt(rightToLeft))
                .Append("\" workbookViewId=\"0\">\n"u8);

            if (splitRow > 0 || splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
            _customWriter.Append("</sheetView></sheetViews>\n"u8);
            WriteColumns(columns);
            _customWriter.Append("<sheetData>\n"u8);
        }

        public void Commit()
        {
            _customWriter.FlushTo(_stream);
        }

        public void Dispose()
        {
            CloseLastRow();
            _customWriter.Append("</sheetData>\n"u8);
            WriteSheetProtection();
            WriteAutoFilter();
            WriteMergedCells();
            WriteDataValidations();
            WriteHeaderFooter();
            WritePageBreaks();
            _customWriter.Append("</worksheet>\n"u8);
            _customWriter.FlushTo(_stream);
            _stream.Dispose();
        }

        public void BeginRow(double? height, bool hidden, XlsxStyle style)
        {
            CloseLastRow();
            if (CurrentRowNumber == Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + 1} attempted)");
            CurrentRowNumber++;
            _stringedCurrentRowNumberLength = 0;
            CurrentColumnNumber = 1;
            _customWriter.Append("<row"u8);
            if (_requireCellReferences || _needsRef)
            {
                _customWriter.Append(" r=\""u8);
                WriteCurrentRowNumber();
                _customWriter.Append("\""u8);
                _needsRef = false;
            }
            if (height.HasValue)
                _customWriter.Append(" ht=\""u8).Append(height.Value).Append("\" customHeight=\"1\""u8);
            if (hidden)
                _customWriter.Append(" hidden=\"1\""u8);
            if (style != null)
                _customWriter.Append(" s=\""u8).Append(_stylesheet.ResolveStyleId(style)).Append("\" customFormat=\"1\""u8);
            _customWriter.Append(">\n"u8);
        }

        public void SkipRows(int rowCount)
        {
            CloseLastRow();
            _needsRef = true;
            if (CurrentRowNumber + rowCount > Limits.MaxRowCount)
                throw new InvalidOperationException($"A worksheet can contain at most {Limits.MaxRowCount} rows ({CurrentRowNumber + rowCount} attempted)");
            CurrentRowNumber += rowCount;
        }

        public void SkipColumns(int columnCount)
        {
            EnsureRow();
            _needsRef = true;
            CurrentColumnNumber += columnCount;
        }


        public void Write(XlsxStyle style, int repeatCount)
        {
            EnsureRow();
            var styleId = _stylesheet.ResolveStyleId(style);
            for (var i = 0; i < repeatCount; i++)
            {
                // <c r="{0}{1}" s="{2}"/>
                _customWriter.Append("<c"u8);
                WriteCellRef();
                WriteStyle(styleId);
                _customWriter.Append("/>\n"u8);
                CurrentColumnNumber++;
            }
        }

        public void Write(string value, XlsxStyle style)
        {
            if (value == null)
            {
                Write(style, 1);
                return;
            }
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="inlineStr"><is><t xml:space="preserve">{3}</t></is></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter
                .Append(" t=\"inlineStr\"><is><t"u8)
                .AddSpacePreserveIfNeeded(value)
                .Append(">"u8)
                .AppendEscapedXmlText(value, _skipInvalidCharacters)
                .Append("</t></is></c>\n"u8);
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter.Append("><v>"u8).Append(value).Append("</v></c>\n"u8);
            CurrentColumnNumber++;
        }

        public void Write(decimal value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter.Append("><v>"u8)
                .Append(value)
                .Append("</v></c>\n"u8);
            CurrentColumnNumber++;
        }

        public void Write(int value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}"><v>{3}</v></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter.Append("><v>"u8).Append(value).Append("</v></c>\n"u8);
            CurrentColumnNumber++;
        }

        public void Write(bool value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="b"><v>{3}</v></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter
                .Append(" t=\"b\"><v>"u8)
                .Append(Util.BoolToInt(value))
                .Append("</v></c>\n"u8);
            CurrentColumnNumber++;
        }

        public void WriteFormula(string formula, XlsxStyle style, IConvertible result)
        {
            // <c r="{0}{1}" s="{2}" t="str"><f>{3}</f><v>{4}</v></c>
            EnsureRow();
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter
                .Append(" t=\"str\"><f>"u8)
                .AppendEscapedXmlText(formula, _skipInvalidCharacters)
                .Append("</f>"u8);
            if (result != null)
                _customWriter
                    .Append("<v>"u8)
                    .AppendEscapedXmlText(result.ToString(CultureInfo.InvariantCulture), _skipInvalidCharacters)
                    .Append("</v>"u8);
            _customWriter.Append("</c>\n"u8);
            CurrentColumnNumber++;
        }

        public void WriteSharedString(string value, XlsxStyle style)
        {
            EnsureRow();
            // <c r="{0}{1}" s="{2}" t="s"><v>{3}</v></c>
            _customWriter.Append("<c"u8);
            WriteCellRef();
            WriteStyle(style);
            _customWriter
                .Append(" t=\"s\"><v>"u8)
                .Append(_sharedStringTable.ResolveStringId(value))
                .Append("</v></c>\n"u8);
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

        public void AddRowPageBreakBefore(int rowNumber)
        {
            if (rowNumber <= 1 || rowNumber > Limits.MaxRowCount)
                throw new ArgumentOutOfRangeException(nameof(rowNumber));
            _pageBreakRowNumbers.Add(rowNumber - 1);
        }

        public void AddColumnPageBreakBefore(int columnNumber)
        {
            if (columnNumber <= 1 || columnNumber > Limits.MaxColumnCount)
                throw new ArgumentOutOfRangeException(nameof(columnNumber));
            _pageBreakColumnNumbers.Add(columnNumber - 1);
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
            if (sheetProtection.Password.Length < Limits.MinSheetProtectionPasswordLength || sheetProtection.Password.Length > Limits.MaxSheetProtectionPasswordLength)
                throw new ArgumentException("Invalid password length");
            _sheetProtection = sheetProtection;
        }

        public void SetHeaderFooter(XlsxHeaderFooter headerFooter)
        {
            _headerFooter = headerFooter;
        }

        private void WriteCellRef()
        {
            if (_requireCellReferences || _needsRef)
            {
                _customWriter.Append(" r=\""u8);
                _customWriter.Append(Util.GetUtf8ColumnName(CurrentColumnNumber));
                WriteCurrentRowNumber();
                _customWriter.Append("\""u8);
                _needsRef = false;
            }
        }

        private void WriteCurrentRowNumber()
        {
            if (_stringedCurrentRowNumberLength == 0)
                _stringedCurrentRowNumberLength = _customWriter.GetUtf8Bytes(CurrentRowNumber, _stringedCurrentRowNumber);
            _customWriter.Append(new ReadOnlySpan<byte>(_stringedCurrentRowNumber, 0, _stringedCurrentRowNumberLength));
        }

        private void WriteStyle(int styleId)
        {
            if (styleId != 0)
                _customWriter.Append(" s=\""u8).Append(styleId).Append("\""u8);
        }

        private void WriteStyle(XlsxStyle style)
        {
            WriteStyle(_stylesheet.ResolveStyleId(style));
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
                _customWriter.Append("</row>\n"u8);
                CurrentColumnNumber = 0;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{Util.GetColumnName(fromColumn + 1)}{fromRow + 1}";
            if (fromRow > 0 && fromColumn > 0)
            {
                _customWriter
                    .Append("<pane xSplit=\""u8)
                    .Append(fromColumn)
                    .Append("\" ySplit=\""u8)
                    .Append(fromRow)
                    .Append("\" topLeftCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" activePane=\"bottomRight\" state=\"frozen\"/><selection pane=\"bottomRight\" activeCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" sqref=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\"/>\n"u8);
            }
            else if (fromRow > 0)
            {
                _customWriter
                    .Append("<pane ySplit=\""u8)
                    .Append(fromRow)
                    .Append("\" topLeftCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" activePane=\"bottomLeft\" state=\"frozen\"/><selection pane=\"bottomLeft\" activeCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" sqref=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\"/>\n"u8);
            }
            else if (fromColumn > 0)
            {
                _customWriter
                    .Append("<pane xSplit=\""u8)
                    .Append(fromColumn)
                    .Append("\" topLeftCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" activePane=\"topRight\" state=\"frozen\"/><selection pane=\"topRight\" activeCell=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\" sqref=\""u8)
                    .AppendEscapedXmlText(topLeftCell, false)
                    .Append("\"/>\n"u8);
            }
        }

        private void WriteColumns(IEnumerable<XlsxColumn> columns)
        {
            var columnIndex = 1;
            var colsWritten = false;
            foreach (var column in columns)
            {
                if (column.Hidden || column.Style != null || column.Width.HasValue)
                {
                    if (!colsWritten)
                    {
                        _customWriter.Append("<cols>"u8);
                        colsWritten = true;
                    }
                    _customWriter.Append("<col min=\""u8).Append(columnIndex).Append("\" max=\""u8).Append(columnIndex + column.Count - 1).Append("\""u8);
                    if (column.Width.HasValue) _customWriter.Append(" width=\""u8).Append(column.Width.Value).Append("\""u8);
                    if (column.Hidden) _customWriter.Append(" hidden=\"1\""u8);
                    if (column.Width.HasValue) _customWriter.Append(" customWidth=\"1\""u8);
                    if (column.Style != null) _customWriter.Append(" style=\""u8).Append(_stylesheet.ResolveStyleId(column.Style)).Append("\""u8);
                    _customWriter.Append("/>\n"u8);
                }
                columnIndex += column.Count;
            }
            if (colsWritten)
                _customWriter.Append("</cols>\n"u8);
        }

        private void WriteAutoFilter()
        {
            if (_autoFilterRef != null)
                _customWriter.Append("<autoFilter ref=\""u8).AppendEscapedXmlAttribute(_autoFilterRef, false).Append("\"/>\n"u8);
        }

        private void WriteMergedCells()
        {
            if (!_mergedCellRefs.Any())
                return;
            _customWriter.Append("<mergeCells count=\""u8).Append(_mergedCellRefs.Count).Append("\">\n"u8);
            foreach (var mergedCell in _mergedCellRefs)
                _customWriter.Append("<mergeCell ref=\""u8).AppendEscapedXmlAttribute(mergedCell, false).Append("\"/>\n"u8);
            _customWriter.Append("</mergeCells>\n"u8);
        }

        private void WriteDataValidations()
        {
            if (!_cellRefsByDataValidation.Any())
                return;
            _customWriter.Append("<dataValidations count=\""u8).Append(_cellRefsByDataValidation.Count).Append("\">\n"u8);
            foreach (var kvp in _cellRefsByDataValidation)
            {
                _customWriter.Append("<dataValidation sqref=\""u8)
                    .AppendEscapedXmlAttribute(string.Join(" ", kvp.Value.Distinct()), false)
                    .Append("\" allowBlank=\""u8)
                    .Append(Util.BoolToInt(kvp.Key.AllowBlank))
                    .Append("\""u8);
                if (kvp.Key.Error != null)
                    _customWriter.Append(" error=\""u8).AppendEscapedXmlAttribute(kvp.Key.Error, _skipInvalidCharacters).Append("\""u8);
                if (kvp.Key.ErrorStyleValue.HasValue)
                    _customWriter.Append(" errorStyle=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(kvp.Key.ErrorStyleValue), false).Append("\""u8);
                if (kvp.Key.ErrorTitle != null)
                    _customWriter.Append(" errorTitle=\""u8).AppendEscapedXmlAttribute(kvp.Key.ErrorTitle, _skipInvalidCharacters).Append("\""u8);
                if (kvp.Key.OperatorValue.HasValue)
                    _customWriter.Append(" operator=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(kvp.Key.OperatorValue), false).Append("\""u8);
                if (kvp.Key.Prompt != null)
                    _customWriter.Append(" prompt=\""u8).AppendEscapedXmlAttribute(kvp.Key.Prompt, _skipInvalidCharacters).Append("\""u8);
                if (kvp.Key.PromptTitle != null)
                    _customWriter.Append(" promptTitle=\""u8).AppendEscapedXmlAttribute(kvp.Key.PromptTitle, _skipInvalidCharacters).Append("\""u8);
                if (kvp.Key.ShowDropDown)
                    _customWriter.Append(" showDropDown=\"1\""u8);
                if (kvp.Key.ShowErrorMessage)
                    _customWriter.Append(" showErrorMessage=\"1\""u8);
                if (kvp.Key.ShowInputMessage)
                    _customWriter.Append(" showInputMessage=\"1\""u8);
                if (kvp.Key.ValidationTypeValue.HasValue)
                    _customWriter.Append(" type=\""u8).AppendEscapedXmlAttribute(Util.EnumToAttributeValue(kvp.Key.ValidationTypeValue), false).Append("\""u8);
                _customWriter.Append(">"u8);
                if (kvp.Key.Formula1 != null)
                    _customWriter.Append("<formula1>"u8).AppendEscapedXmlText(kvp.Key.Formula1, _skipInvalidCharacters).Append("</formula1>"u8);
                if (kvp.Key.Formula2 != null)
                    _customWriter.Append("<formula2>"u8).AppendEscapedXmlText(kvp.Key.Formula2, _skipInvalidCharacters).Append("</formula2>"u8);
                _customWriter.Append("</dataValidation>\n"u8);
            }
            _customWriter.Append("</dataValidations>\n"u8);
        }

        private void WriteSheetProtection()
        {
            if (_sheetProtection == null)
                return;
            const int spinCount = 100000;
            var saltValue = Guid.NewGuid().ToByteArray();
            var hash = Util.ComputePasswordHash(_sheetProtection.Password, saltValue, spinCount);
            _customWriter
                .Append("<sheetProtection algorithmName=\"SHA-512\" hashValue=\""u8)
                .AppendEscapedXmlAttribute(Convert.ToBase64String(hash), false)
                .Append("\" saltValue=\""u8)
                .AppendEscapedXmlAttribute(Convert.ToBase64String(saltValue), false)
                .Append("\" spinCount=\""u8)
                .Append(spinCount)
                .Append("\""u8);
            if (_sheetProtection.Sheet) _customWriter.Append(" sheet=\"1\""u8);
            if (_sheetProtection.Objects) _customWriter.Append(" objects=\"1\""u8);
            if (_sheetProtection.Scenarios) _customWriter.Append(" scenarios=\"1\""u8);
            if (!_sheetProtection.FormatCells) _customWriter.Append(" formatCells=\"0\""u8);
            if (!_sheetProtection.FormatColumns) _customWriter.Append(" formatColumns=\"0\""u8);
            if (!_sheetProtection.FormatRows) _customWriter.Append(" formatRows=\"0\""u8);
            if (!_sheetProtection.InsertColumns) _customWriter.Append(" insertColumns=\"0\""u8);
            if (!_sheetProtection.InsertRows) _customWriter.Append(" insertRows=\"0\""u8);
            if (!_sheetProtection.InsertHyperlinks) _customWriter.Append(" insertHyperlinks=\"0\""u8);
            if (!_sheetProtection.DeleteColumns) _customWriter.Append(" deleteColumns=\"0\""u8);
            if (!_sheetProtection.DeleteRows) _customWriter.Append(" deleteRows=\"0\""u8);
            if (_sheetProtection.SelectLockedCells) _customWriter.Append(" selectLockedCells=\"1\""u8);
            if (!_sheetProtection.Sort) _customWriter.Append(" sort=\"0\""u8);
            if (!_sheetProtection.AutoFilter) _customWriter.Append(" autoFilter=\"0\""u8);
            if (!_sheetProtection.PivotTables) _customWriter.Append(" pivotTables=\"0\""u8);
            if (_sheetProtection.SelectUnlockedCells) _customWriter.Append(" selectUnlockedCells=\"1\""u8);
            _customWriter.Append("/>\n"u8);
        }

        private void WriteHeaderFooter()
        {
            if (_headerFooter == null)
                return;
            var differentFirst = _headerFooter.FirstHeader != null || _headerFooter.FirstFooter != null;
            var differentOddEven = _headerFooter.EvenHeader != null || _headerFooter.EvenFooter != null;
            _customWriter
                .Append("<headerFooter alignWithMargins=\""u8)
                .Append(Util.BoolToInt(_headerFooter.AlignWithMargins))
                .Append("\" differentFirst=\""u8)
                .Append(Util.BoolToInt(differentFirst))
                .Append("\" differentOddEven=\""u8)
                .Append(Util.BoolToInt(differentOddEven))
                .Append("\" scaleWithDoc=\""u8)
                .Append(Util.BoolToInt(_headerFooter.ScaleWithDoc))
                .Append("\">\n"u8);
            if (_headerFooter.OddHeader != null)
                _customWriter.Append("<oddHeader>"u8).AppendEscapedXmlText(_headerFooter.OddHeader, _skipInvalidCharacters).Append("</oddHeader>\n"u8);
            if (_headerFooter.OddFooter != null)
                _customWriter.Append("<oddFooter>"u8).AppendEscapedXmlText(_headerFooter.OddFooter, _skipInvalidCharacters).Append("</oddFooter>\n"u8);
            if (_headerFooter.EvenHeader != null)
                _customWriter.Append("<evenHeader>"u8).AppendEscapedXmlText(_headerFooter.EvenHeader, _skipInvalidCharacters).Append("</evenHeader>\n"u8);
            if (_headerFooter.EvenFooter != null)
                _customWriter.Append("<evenFooter>"u8).AppendEscapedXmlText(_headerFooter.EvenFooter, _skipInvalidCharacters).Append("</evenFooter>\n"u8);
            if (_headerFooter.FirstHeader != null)
                _customWriter.Append("<firstHeader>"u8).AppendEscapedXmlText(_headerFooter.FirstHeader, _skipInvalidCharacters).Append("</firstHeader>\n"u8);
            if (_headerFooter.FirstFooter != null)
                _customWriter.Append("<firstFooter>"u8).AppendEscapedXmlText(_headerFooter.FirstFooter, _skipInvalidCharacters).Append("</firstFooter>\n"u8);
            _customWriter.Append("</headerFooter>\n"u8);
        }

        private void WritePageBreaks()
        {
            if (_pageBreakRowNumbers.Count > 0)
            {
                _customWriter.Append("<rowBreaks count=\""u8).Append(_pageBreakRowNumbers.Count).Append("\" manualBreakCount=\""u8).Append(_pageBreakRowNumbers.Count).Append("\">\n"u8);
                foreach (var i in _pageBreakRowNumbers.OrderBy(r => r))
                    _customWriter.Append("<brk id=\""u8).Append(i).Append("\" max=\""u8).Append(Limits.MaxColumnCount).Append("\" man=\"1\"/>\n"u8);
                _customWriter.Append("</rowBreaks>\n"u8);
            }
            if (_pageBreakColumnNumbers.Count > 0)
            {
                _customWriter.Append("<colBreaks count=\""u8).Append(_pageBreakColumnNumbers.Count).Append("\" manualBreakCount=\""u8).Append(_pageBreakColumnNumbers.Count).Append("\">\n"u8);
                foreach (var i in _pageBreakColumnNumbers.OrderBy(c => c))
                    _customWriter.Append("<brk id=\""u8).Append(i).Append("\" max=\""u8).Append(Limits.MaxRowCount).Append("\" man=\"1\"/>\n"u8);
                _customWriter.Append("</colBreaks>\n"u8);
            }
        }
    }
}