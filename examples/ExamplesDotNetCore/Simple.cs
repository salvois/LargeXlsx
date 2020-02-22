using System.IO;
using LargeXlsx;

namespace ExamplesDotNetCore
{
    public static class Simple
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Simple)}.xlsx", FileMode.Create))
            using (var largeXlsxWriter = new XlsxWriter(stream))
            {
                var whiteFont = largeXlsxWriter.Stylesheet.CreateFont("Segoe UI", 9, "ffffff", bold: true);
                var blueFill = largeXlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = largeXlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);

                largeXlsxWriter.BeginWorksheet("Sheet1")
                    .BeginRow().Write("Col1", headerStyle).Write("Col2", headerStyle).Write("Col3", headerStyle)
                    .BeginRow().Write("Row2").Write(42).Write(-1)
                    .BeginRow().Write("Row3").SkipColumns(1).Write(1234)
                    .SkipRows(2)
                    .BeginRow().AddMergedCell(1, 2).Write("Row6").SkipColumns(1).Write(3.14159265359);
            }
        }
    }
}
