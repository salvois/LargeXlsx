using System;
using System.Diagnostics;
using System.IO;
using LargeXlsx;

namespace ExamplesDotNetCore
{
    public static class Large
    {
        public static void Run()
        {
            var stopwatch = Stopwatch.StartNew();
            using (var stream = new FileStream($"{nameof(Large)}.xlsx", FileMode.Create))
            using (var largeXlsxWriter = new LargeXlsxWriter(stream))
            {
                var whiteFontId = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff");
                var blueFillId = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyleId = largeXlsxWriter.Stylesheet.CreateStyle(whiteFontId, blueFillId, LargeXlsxStylesheet.GeneralNumberFormatId, LargeXlsxStylesheet.NoBorderId);
                var numberStyleId = largeXlsxWriter.Stylesheet.CreateStyle(LargeXlsxStylesheet.DefaultFontId, LargeXlsxStylesheet.NoFillId, LargeXlsxStylesheet.TwoDecimalExcelNumberFormatId, LargeXlsxStylesheet.NoBorderId);

                largeXlsxWriter.BeginSheet("Sheet1", 1, 1);
                largeXlsxWriter.BeginRow();
                for (var j = 0; j < 180; j++)
                    largeXlsxWriter.WriteInlineStringCell($"Column {j}", headerStyleId);
                for (var i = 0; i < 50000; i++)
                {
                    largeXlsxWriter.BeginRow().WriteInlineStringCell($"Row {i}");
                    for (var j = 1; j < 180; j++)
                        largeXlsxWriter.WriteNumericCell(i * 1000 + j, numberStyleId);
                }
            }
            stopwatch.Stop();
            Console.WriteLine($"Completed in {stopwatch.ElapsedMilliseconds} ms. Press any key...");
            Console.ReadKey();
        }
    }
}