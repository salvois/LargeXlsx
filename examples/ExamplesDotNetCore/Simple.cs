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
                var whiteFont = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff", bold: true);
                var blueFill = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = largeXlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, LargeXlsxStylesheet.GeneralNumberFormat, LargeXlsxStylesheet.NoBorder);

                largeXlsxWriter.BeginSheet("Sheet1")
                    .BeginRow().WriteInlineString("Col1", headerStyle).WriteInlineString("Col2", headerStyle).WriteInlineString("Col3", headerStyle)
                    .BeginRow().WriteInlineString("Row2").Write(42).Write(-1)
                    .BeginRow().WriteInlineString("Row3").SkipColumns(1).Write(1234)
                    .SkipRows(2)
                    .BeginRow().WriteInlineString("Row6").AddMergedCell(1, 2).SkipColumns(1).Write(3.14159265359);
            }
        }
    }
}
