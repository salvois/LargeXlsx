using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using LargeXlsx;

namespace Examples
{
    public static class StyledLarge
    {
        public static void Run()
        {
            var rnd = new Random();
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(StyledLarge)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var whiteFont = new XlsxFont("Calibri", 11, Color.White, bold: true);
                var blueFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86));
                var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                var cellStyles = Enumerable.Repeat(0, 100)
                    .Select(_ =>
                    {
                        var color = Color.FromArgb(rnd.Next(256), rnd.Next(256), rnd.Next(256));
                        var fill = new XlsxFill(XlsxFill.Pattern.Solid, color);
                        return new XlsxStyle(XlsxFont.Default, fill, XlsxBorder.None, XlsxNumberFormat.General);
                    })
                    .ToList();

                xlsxWriter.BeginWorksheet("Sheet1", 1, 1);
                xlsxWriter.BeginRow();
                for (var j = 0; j < 180; j++)
                    xlsxWriter.Write($"Column {j}", headerStyle);
                var cellStyleIndex = 0;
                for (var i = 0; i < 50000; i++)
                {
                    xlsxWriter.BeginRow().Write($"Row {i}");
                    for (var j = 1; j < 180; j++)
                    {
                        xlsxWriter.Write(i * 1000 + j, cellStyles[cellStyleIndex]);
                        cellStyleIndex = (cellStyleIndex + 1) % cellStyles.Count;
                    }
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"Completed in {stopwatch.ElapsedMilliseconds} ms. Press any key...");
            Console.ReadKey();
        }
    }
}