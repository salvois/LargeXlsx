using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class Large
    {
        private const int RowCount = 50000;
        private const int ColumnCount = 180;

        public static void Run()
        {
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(Large)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var whiteFont = new XlsxFont("Calibri", 11, Color.White, bold: true);
                var blueFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86));
                var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                var numberStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.ThousandTwoDecimal);

                xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
                xlsxWriter.BeginRow();
                for (var j = 0; j < ColumnCount; j++)
                    xlsxWriter.Write($"Column {j}", headerStyle);
                for (var i = 0; i < RowCount; i++)
                {
                    xlsxWriter.BeginRow().Write($"Row {i}");
                    for (var j = 1; j < ColumnCount; j++)
                        xlsxWriter.Write(i * 1000 + j, numberStyle);
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"{nameof(Large)} completed {RowCount} rows and {ColumnCount} columns in {stopwatch.ElapsedMilliseconds} ms.");
        }
    }
}