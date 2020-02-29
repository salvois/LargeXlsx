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
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LargeXlsx
{
    internal class XlsxWorksheet : IDisposable
    {
        private static readonly Row RowElement = new Row();
        private static readonly Cell CellElement = new Cell();
        private static readonly OpenXmlAttribute InlineStringTypeAttribute = new OpenXmlAttribute("t", null, "inlineStr");
        private static readonly OpenXmlAttribute NumberTypeAttribute = new OpenXmlAttribute("t", null, "n");
        private readonly List<string> _mergedCells;
        private readonly OpenXmlWriter _worksheetWriter;

        public string Name { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }
        public WorksheetPart WorksheetPart { get; }

        public XlsxWorksheet(SpreadsheetDocument document, string name, int splitRow, int splitColumn)
        {
            Name = name;
            CurrentRowNumber = 0;
            CurrentColumnNumber = 0;
            _mergedCells = new List<string>();

            WorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
            _worksheetWriter = OpenXmlWriter.Create(WorksheetPart);
            _worksheetWriter.WriteStartElement(new Worksheet());
            if (splitRow > 0 && splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
            _worksheetWriter.WriteStartElement(new SheetData());
        }

        public void Dispose()
        {
            CloseLastRow();
            _worksheetWriter.WriteEndElement(); // sheetdata
            WriteMergedCells();
            _worksheetWriter.WriteEndElement(); // worksheet
            _worksheetWriter.Close();
        }

        public void BeginRow()
        {
            CloseLastRow();
            CurrentRowNumber++;
            CurrentColumnNumber = 1;
            _worksheetWriter.WriteStartElement(RowElement, new[]
            {
                new OpenXmlAttribute("r", null, CurrentRowNumber.ToString())
            });
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
            _worksheetWriter.WriteStartElement(CellElement, new[]
            {
                new OpenXmlAttribute("r", null, $"{GetColumnName(CurrentColumnNumber)}{CurrentRowNumber}"),
                new OpenXmlAttribute("s", null, style.Id.ToString())
            });
            _worksheetWriter.WriteEndElement();
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
            _worksheetWriter.WriteStartElement(CellElement, new[]
            {
                new OpenXmlAttribute("r", null, $"{GetColumnName(CurrentColumnNumber)}{CurrentRowNumber}"),
                new OpenXmlAttribute("s", null, style.Id.ToString()),
                InlineStringTypeAttribute
            });
            _worksheetWriter.WriteStartElement(new InlineString());
            _worksheetWriter.WriteElement(new Text(value));
            _worksheetWriter.WriteEndElement();
            _worksheetWriter.WriteEndElement();
            CurrentColumnNumber++;
        }

        public void Write(double value, XlsxStyle style)
        {
            EnsureRow();
            _worksheetWriter.WriteStartElement(CellElement, new[]
            {
                new OpenXmlAttribute("r", null, $"{GetColumnName(CurrentColumnNumber)}{CurrentRowNumber}"),
                new OpenXmlAttribute("s", null, style.Id.ToString()),
                NumberTypeAttribute
            });
            _worksheetWriter.WriteElement(new CellValue(value.ToString(CultureInfo.InvariantCulture)));
            _worksheetWriter.WriteEndElement();
            CurrentColumnNumber++;
        }

        public void AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
        {
            var toRow = fromRow + rowCount - 1;
            var fromColumnName = GetColumnName(fromColumn);
            var toColumnName = GetColumnName(fromColumn + columnCount - 1);
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
                _worksheetWriter.WriteEndElement();
                CurrentColumnNumber = 0;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{GetColumnName(fromColumn + 1)}{fromRow + 1}";

            _worksheetWriter.WriteStartElement(new SheetViews());
            _worksheetWriter.WriteStartElement(new SheetView(), new[]
            {
                new OpenXmlAttribute("tabSelected", null, "1"),
                new OpenXmlAttribute("workbookViewId", null, "0"),
            });

            _worksheetWriter.WriteStartElement(new Pane(), new[]
            {
                new OpenXmlAttribute("xSplit", null, fromColumn.ToString()),
                new OpenXmlAttribute("ySplit", null, fromRow.ToString()),
                new OpenXmlAttribute("topLeftCell", null, topLeftCell),
                new OpenXmlAttribute("activePane", null, "bottomRight"),
                new OpenXmlAttribute("state", null, "frozen"),
            });
            _worksheetWriter.WriteEndElement();

            _worksheetWriter.WriteStartElement(new Selection(), new[]
            {
                new OpenXmlAttribute("pane", null, "bottomRight"),
                new OpenXmlAttribute("activeCell", null, topLeftCell),
                new OpenXmlAttribute("sqref", null, topLeftCell),
            });
            _worksheetWriter.WriteEndElement();

            _worksheetWriter.WriteEndElement();
            _worksheetWriter.WriteEndElement();
        }

        private void WriteMergedCells()
        {
            if (!_mergedCells.Any()) return;

            _worksheetWriter.WriteStartElement(new MergeCells(), new[]
            {
                new OpenXmlAttribute("count", null, _mergedCells.Count.ToString()),
            });
            foreach (var mergedCell in _mergedCells)
            {
                _worksheetWriter.WriteStartElement(new MergeCell(), new[]
                {
                    new OpenXmlAttribute("ref", null, mergedCell),
                });
                _worksheetWriter.WriteEndElement();
            }
            _worksheetWriter.WriteEndElement();
        }

        private static string GetColumnName(int columnIndex)
        {
            var columnName = new StringBuilder(3);
            while (true)
            {
                if (columnIndex > 26)
                {
                    columnIndex = Math.DivRem(columnIndex - 1, 26, out var rem);
                    columnName.Insert(0, (char)('A' + rem));
                }
                else
                {
                    columnName.Insert(0, (char)('A' + columnIndex - 1));
                    return columnName.ToString();
                }
            }
        }
    }
}