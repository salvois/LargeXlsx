using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using SharpCompress.Writers.Zip;

namespace LargeXlsx
{
    internal class XlsxSheet2 : IDisposable
    {
        private readonly Stream _stream;
        private readonly StreamWriter _streamWriter;
        private readonly List<string> _mergedCells;

        public string Name { get; }
        public int CurrentRowNumber { get; private set; }
        public int CurrentColumnNumber { get; private set; }

        public XlsxSheet2(ZipWriter zipWriter, string name, int splitRow, int splitColumn)
        {
            Name = name;
            CurrentRowNumber = 0;
            CurrentColumnNumber = -1;
            _mergedCells = new List<string>();

            _stream = zipWriter.WriteToStream("xl/worksheets/sheet1.xml", new ZipWriterEntryOptions());
            _streamWriter = new InvariantCultureStreamWriter(_stream);

            _streamWriter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            if (splitRow > 0 && splitColumn > 0)
                FreezePanes(splitRow, splitColumn);
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
            CurrentColumnNumber = 0;
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

        public void WriteInlineStringCell(string value, XlsxStyle style)
        {
            EnsureRow();
            CurrentColumnNumber++;
            _streamWriter.Write("<c r=\"{0}{1}\" s=\"0\" t=\"inlineStr\"><is><t>{2}</t></is></c>", GetColumnName(CurrentColumnNumber), CurrentRowNumber, value);
        }

        public void WriteNumericCell(double value, XlsxStyle style)
        {
            EnsureRow();
            CurrentColumnNumber++;
            _streamWriter.Write("<c r=\"{0}{1}\" s=\"0\" t=\"n\"><v>{2}</v></c>", GetColumnName(CurrentColumnNumber), CurrentRowNumber, value);
        }

        public void AddMergedCell(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            var fromColumnName = GetColumnName(fromColumn);
            var toColumnName = GetColumnName(toColumn);
            _mergedCells.Add($"{fromColumnName}{fromRow}:{toColumnName}{toRow}");
        }

        private void EnsureRow()
        {
            if (CurrentColumnNumber < 0)
                throw new InvalidOperationException($"{nameof(BeginRow)} not called");
        }

        private void CloseLastRow()
        {
            if (CurrentColumnNumber >= 0)
            {
                _streamWriter.Write("</row>");
                CurrentColumnNumber = -1;
            }
        }

        private void FreezePanes(int fromRow, int fromColumn)
        {
            var topLeftCell = $"{GetColumnName(fromColumn + 1)}{fromRow + 1}";
            _streamWriter.Write("<sheetViews>"
                                + "<sheetView tabSelected=\"1\" workbookViewId=\"0\">"
                                + "<pane xSplit=\"{0}\" ySplit=\"{1}\" topLeftCell=\"{2}\" activePane=\"bottomRight\" state=\"frozen\"/>"
                                + "<selection pane=\"bottomRight\" activeCell=\"{2}\" sqref=\"{2}\"/>"
                                + "</sheetView>"
                                + "</sheetViews>",
                fromColumn, fromRow, topLeftCell);
        }

        private void WriteMergedCells()
        {
            if (!_mergedCells.Any()) return;

            _streamWriter.Write("<mergeCells count=\"{0}\">", _mergedCells.Count);
            foreach (var mergedCell in _mergedCells)
                _streamWriter.Write("<mergeCell ref=\"{0}\"/>", mergedCell);
            _streamWriter.Write("</mergeCells>");
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