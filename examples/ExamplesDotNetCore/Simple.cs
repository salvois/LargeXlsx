using System.IO;
using LargeXlsx;

namespace ExamplesDotNetCore
{
    public static class Simple
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Simple)}.xlsx", FileMode.Create))
            using (var largeXlsxWriter = new LargeXlsxWriter(stream))
            {
                var whiteFontId = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff");
                var blueFillId = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyleId = largeXlsxWriter.Stylesheet.CreateStyle(whiteFontId, blueFillId, LargeXlsxStylesheet.GeneralNumberFormatId, LargeXlsxStylesheet.NoBorderId);

                largeXlsxWriter.BeginSheet("Sheet1")
                    .BeginRow().WriteInlineStringCell("Col1", headerStyleId).WriteInlineStringCell("Col2", headerStyleId).WriteInlineStringCell("Col3", headerStyleId)
                    .BeginRow().WriteInlineStringCell("Row2").WriteNumericCell(42).WriteNumericCell(-1)
                    .BeginRow().WriteInlineStringCell("Row3").SkipColumns(1).WriteNumericCell(1234)
                    .SkipRows(2)
                    .BeginRow().WriteInlineStringCell("Row6").AddMergedCell(6, 1, 6, 2).SkipColumns(1).WriteNumericCell(3.14159265359);
            }
        }
    }
}
