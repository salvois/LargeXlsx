/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2019 Salvatore ISAJA. All rights reserved.

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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LargeXlsx
{
    public class LargeXlsxWriter : IDisposable
    {
        private readonly SpreadsheetDocument _document;
        private readonly List<LargeXlsxSheet> _largeXlsxSheets;
        private LargeXlsxSheet _currentSheet;

        public LargeXlsxStylesheet Stylesheet { get; }

        public LargeXlsxWriter(Stream stream)
        {
            _largeXlsxSheets = new List<LargeXlsxSheet>();
            Stylesheet = new LargeXlsxStylesheet();
            _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public void Dispose()
        {
            _currentSheet?.Dispose();
            Stylesheet.Save(_document);
            _document.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };
            _document.WorkbookPart.Workbook.Sheets.Append(_largeXlsxSheets.Select((largeXlsxSheet, index) => new Sheet
            {
                Name = largeXlsxSheet.Name,
                SheetId = (uint)(index + 1),
                Id = _document.WorkbookPart.GetIdOfPart(largeXlsxSheet.WorksheetPart)
            }));
            _document.Close();
        }

        public LargeXlsxWriter BeginSheet(string name, int splitRow = 0, int splitColumn = 0)
        {
            _currentSheet?.Dispose();
            _currentSheet = new LargeXlsxSheet(_document, name, splitRow, splitColumn);
            _largeXlsxSheets.Add(_currentSheet);
            return this;
        }

        public LargeXlsxWriter SkipRows(int rowCount)
        {
            EnsureSheet();
            _currentSheet.SkipRows(rowCount);
            return this;
        }

        public LargeXlsxWriter BeginRow()
        {
            EnsureSheet();
            _currentSheet.BeginRow();
            return this;
        }

        public LargeXlsxWriter SkipColumns(int columnCount)
        {
            EnsureSheet();
            _currentSheet.SkipColumns(columnCount);
            return this;
        }

        public LargeXlsxWriter WriteInlineStringCell(string value, uint styleId = 0)
        {
            EnsureSheet();
            _currentSheet.WriteInlineStringCell(value, styleId);
            return this;
        }

        public LargeXlsxWriter WriteNumericCell(double value, uint styleId = 0)
        {
            EnsureSheet();
            _currentSheet.WriteNumericCell(value, styleId);
            return this;
        }

        public LargeXlsxWriter AddMergedCell(int fromRow, int fromColumn, int toRow, int toColumn)
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
