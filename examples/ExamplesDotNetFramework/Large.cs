using System;
using System.Diagnostics;
using System.IO;
using LargeXlsx;

namespace ExamplesDotNetFramework
{
    public static class Large
    {
        public static void Run()
        {
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(Large)}.xlsx", FileMode.Create))
            using (var largeXlsxWriter = new LargeXlsxWriter(stream))
            {
                var whiteFont = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff", bold: true);
                var blueFill = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = largeXlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, LargeXlsxStylesheet.GeneralNumberFormat, LargeXlsxStylesheet.NoBorder);
                var numberStyle = largeXlsxWriter.Stylesheet.CreateStyle(LargeXlsxStylesheet.DefaultFont, LargeXlsxStylesheet.NoFill, LargeXlsxStylesheet.TwoDecimalExcelNumberFormat, LargeXlsxStylesheet.NoBorder);

                largeXlsxWriter.BeginSheet("Sheet1", 1, 1);
                largeXlsxWriter.BeginRow();
                for (var j = 0; j < 180; j++)
                    largeXlsxWriter.WriteInlineString($"Column {j}", headerStyle);
                for (var i = 0; i < 50000; i++)
                {
                    largeXlsxWriter.BeginRow().WriteInlineString($"Row {i}");
                    for (var j = 1; j < 180; j++)
                        largeXlsxWriter.Write(i * 1000 + j, numberStyle);
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"Completed in {stopwatch.ElapsedMilliseconds} ms. Press any key...");
            Console.ReadKey();
        }
    }
}