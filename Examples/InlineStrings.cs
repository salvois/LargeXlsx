using System;
using System.Diagnostics;
using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class InlineStrings
    {
        private const int RowCount = 1_000_000;

        public static void Run()
        {
            var stopwatch = Stopwatch.StartNew();
            DoRun();
            stopwatch.Stop();
            Console.WriteLine($"{nameof(InlineStrings)} completed in {stopwatch.ElapsedMilliseconds} ms.");
        }

        private static void DoRun()
        {
            using (var stream = new FileStream($"{nameof(InlineStrings)}.xlsx", FileMode.Create, FileAccess.Write))
            {
                using (var xlsxWriter = new XlsxWriter(stream))
                {
                    xlsxWriter.BeginWorksheet("Sheet1");
                    for (var i = 0; i < RowCount; i++)
                    {
                        xlsxWriter.BeginRow()
                            .Write("  Leading spaces")
                            .Write("Trailing spaces   ")
                            .Write("Spaces  in   between");
                    }
                }
            }
        }
    }
}