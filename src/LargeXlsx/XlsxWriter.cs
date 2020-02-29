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
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace LargeXlsx
{
    public class XlsxWriter : IDisposable
    {
        private readonly SpreadsheetDocument _document;
        private readonly List<XlsxWorksheet> _worksheets;
        private XlsxWorksheet _currentWorksheet;

        public XlsxStylesheet Stylesheet { get; }
        public int CurrentRowNumber => _currentWorksheet.CurrentRowNumber;
        public int CurrentColumnNumber => _currentWorksheet.CurrentColumnNumber;

        public XlsxWriter(Stream stream)
        {
            _worksheets = new List<XlsxWorksheet>();
            Stylesheet = new XlsxStylesheet();
            _document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook);
            _document.AddWorkbookPart();
        }

        public void Dispose()
        {
            _currentWorksheet?.Dispose();
            Stylesheet.Save(_document);
            _document.WorkbookPart.Workbook = new Workbook { Sheets = new Sheets() };
            _document.WorkbookPart.Workbook.Sheets.Append(_worksheets.Select((worksheet, index) => new Sheet
            {
                Name = worksheet.Name,
                SheetId = (uint)(index + 1),
                Id = _document.WorkbookPart.GetIdOfPart(worksheet.WorksheetPart)
            }));
            _document.Close();
        }

        public XlsxWriter BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0)
        {
            _currentWorksheet?.Dispose();
            _currentWorksheet = new XlsxWorksheet(_document, name, splitRow, splitColumn);
            _worksheets.Add(_currentWorksheet);
            return this;
        }

        public XlsxWriter SkipRows(int rowCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipRows(rowCount);
            return this;
        }

        public XlsxWriter BeginRow()
        {
            EnsureWorksheet();
            _currentWorksheet.BeginRow();
            return this;
        }

        public XlsxWriter SkipColumns(int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.SkipColumns(columnCount);
            return this;
        }

        public XlsxWriter Write()
        {
            return Write(XlsxStyle.Default);
        }

        public XlsxWriter Write(XlsxStyle style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(style);
            return this;
        }

        public XlsxWriter Write(string value)
        {
            return Write(value, XlsxStyle.Default);
        }

        public XlsxWriter Write(string value, XlsxStyle style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter Write(double value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, XlsxStyle.Default);
            return this;
        }

        public XlsxWriter Write(double value, XlsxStyle style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter Write(decimal value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write((double)value, XlsxStyle.Default);
            return this;
        }

        public XlsxWriter Write(decimal value, XlsxStyle style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write((double)value, style);
            return this;
        }

        public XlsxWriter Write(int value)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, XlsxStyle.Default);
            return this;
        }

        public XlsxWriter Write(int value, XlsxStyle style)
        {
            EnsureWorksheet();
            _currentWorksheet.Write(value, style);
            return this;
        }

        public XlsxWriter AddMergedCell(int rowCount, int columnCount)
        {
            EnsureWorksheet();
            _currentWorksheet.AddMergedCell(_currentWorksheet.CurrentRowNumber, _currentWorksheet.CurrentColumnNumber, rowCount, columnCount);
            return this;
        }

        public XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)
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
