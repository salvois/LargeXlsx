using System.IO;
using LargeXlsx;

namespace ExamplesDotNetFramework
{
    public static class Simple
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Simple)}.xlsx", FileMode.Create))
            using (var largeXlsxWriter = new LargeXlsxWriter(stream))
            {
                var whiteFont = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff");
                var blueFill = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = largeXlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, LargeXlsxStylesheet.GeneralNumberFormat, LargeXlsxStylesheet.NoBorder);

                largeXlsxWriter.BeginSheet("Sheet1")
                    .BeginRow().WriteInlineStringCell("Col1", headerStyle).WriteInlineStringCell("Col2", headerStyle).WriteInlineStringCell("Col3", headerStyle)
                    .BeginRow().WriteInlineStringCell("Row2").WriteNumericCell(42).WriteNumericCell(-1)
                    .BeginRow().WriteInlineStringCell("Row3").SkipColumns(1).WriteNumericCell(1234)
                    .SkipRows(2)
                    .BeginRow().WriteInlineStringCell("Row6").AddMergedCell(6, 1, 6, 2).SkipColumns(1).WriteNumericCell(3.14159265359);
            }
        }
    }
}
