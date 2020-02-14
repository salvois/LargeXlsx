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
                var whiteFont = largeXlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff", bold: true);
                var blueFill = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = largeXlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, LargeXlsxStylesheet.GeneralNumberFormat, LargeXlsxStylesheet.NoBorder);

                largeXlsxWriter.BeginSheet("Sheet1")
                    .BeginRow().Write("Col1", headerStyle).Write("Col2", headerStyle).Write("Col3", headerStyle)
                    .BeginRow().Write(headerStyle).Write("Sub2", headerStyle).Write("Sub3", headerStyle)
                    .BeginRow().Write("Row3").Write(42).Write(-1)
                    .BeginRow().Write("Row4").SkipColumns(1).Write(1234)
                    .SkipRows(2)
                    .BeginRow().Write("Row7").AddMergedCell(1, 2).SkipColumns(1).Write(3.14159265359);
            }
        }
    }
}
