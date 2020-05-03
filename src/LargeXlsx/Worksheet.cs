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
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    internal class Worksheet : IDisposable
    {
        private readonly Stream _stream;
        private readonly StreamWriter _streamWriter;
        private readonly Stylesheet _stylesheet;
        private readonly List<string> _mergedCells;

        public int Id { get; }
        public string Name { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }

        public Worksheet(ZipWriter zipWriter, int id, string name, int splitRow, int splitColumn, Stylesheet stylesheet, IEnumerable<XlsxColumn> columns)
        {
            Id = id;
            Name = name;
            CurrentRowNumber = 0;
            CurrentColumnNumber = 0;
            _stylesheet = stylesheet;
            _mergedCells = new List<string>();
            _stream = zipWriter.WriteToStream($"xl/worksheets/sheet{id}.xml", new ZipWriterEntryOptions());
            _streamWriter = new InvariantCultureStreamWriter(_stream);

            _streamWriter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (splitRow > 0 && splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
            if (columns.Any())
                WriteColumns(columns);
            _streamWriter.Write("<sheetData>");
        }

        public void Dispose()
        {
            CloseLastRow();
            _streamWriter.Write("</sheetData>");
            WriteMergedCells();
            _streamWriter.Write("</worksheet>");
            _streamWriter.Dispose();
            _stream.Dispose();
        }

        public void BeginRow()
        {
            CloseLastRow();
            CurrentRowNumber++;
            CurrentColumnNumber = 1;
            _streamWriter.Write("<row r=\"{0}\">", CurrentRowNumber);
        }

        public void SkipRows(int rowCount)
        {
            CloseLastRow();
            CurrentRowNumber += rowCount;
        }

        public void SkipColumns(int columnCount)
        {
            EnsureRow();
            CurrentColumnNumber += columnCount;
        }

        public void Write(XlsxStyle style)
        {
            EnsureRow();
            var styleId = _stylesheet.ResolveStyleId(style);
            _streamWriter.Write("<c r=\"{0}{1}\" s=\"{2}\"/>", Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, styleId);
            CurrentColumnNumber++;
        }

        public void Write(string value, XlsxStyle style)
        {
            if (value == null)
            {
                Write(style);
                return;
            }

            EnsureRow();
            var escapedValue = Util.EscapeXmlText(value);
            var styleId = _stylesheet.ResolveStyleId(style);
            _streamWriter.Write("<c r=\"{0}{1}\" s=\"{2}\" t=\"inlineStr\"><is><t>{3}</t></is></c>", Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, styleId, escapedValue);
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            EnsureRow();
            var styleId = _stylesheet.ResolveStyleId(style);
            _streamWriter.Write("<c r=\"{0}{1}\" s=\"{2}\" t=\"n\"><v>{3}</v></c>", Util.GetColumnName(CurrentColumnNumber), CurrentRowNumber, styleId, value);
            CurrentColumnNumber++;
        }

        public void AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            if (rowCount < 1 || columnCount < 1)
                throw new ArgumentOutOfRangeException();
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = Util.GetColumnName(fromColumn);
            var toColumnName = Util.GetColumnName(fromColumn + columnCount - 1);
            _mergedCells.Add($"{fromColumnName}{fromRow}:{toColumnName}{toRow}");
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
                _streamWriter.Write("</row>");
                CurrentColumnNumber = 0;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{Util.GetColumnName(fromColumn + 1)}{fromRow + 1}";
            _streamWriter.Write("<sheetViews>"
                                + "<sheetView tabSelected=\"1\" workbookViewId=\"0\">"
                                + "<pane xSplit=\"{0}\" ySplit=\"{1}\" topLeftCell=\"{2}\" activePane=\"bottomRight\" state=\"frozen\"/>"
                                + "<selection pane=\"bottomRight\" activeCell=\"{2}\" sqref=\"{2}\"/>"
                                + "</sheetView>"
                                + "</sheetViews>",
                fromColumn, fromRow, topLeftCell);
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
                    _streamWriter.Write("/>");
                }
                columnIndex += column.Count;
            }
            _streamWriter.Write("</cols>");
        }

        private void WriteMergedCells()
        {
            if (!_mergedCells.Any()) return;

            _streamWriter.Write("<mergeCells count=\"{0}\">", _mergedCells.Count);
            foreach (var mergedCell in _mergedCells)
                _streamWriter.Write("<mergeCell ref=\"{0}\"/>", mergedCell);
            _streamWriter.Write("</mergeCells>");
        }
    }
}