using System.IO;
using LargeXlsx;

namespace ExamplesDotNetCore
{
    public static class Simple
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Simple)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter2(stream))
            {
                var whiteFont = xlsxWriter.Stylesheet.CreateFont("Segoe UI", 9, "ffffff", bold: true);
                var blueFill = xlsxWriter.Stylesheet.CreateSolidFill("004586");
                var headerStyle = xlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, XlsxBorder2.None, XlsxNumberFormat2.General);

                xlsxWriter
                    .BeginWorksheet("Sheet1")
                    .BeginRow().Write("Col1", headerStyle).Write("Col2", headerStyle).Write("Col3", headerStyle)
                    .BeginRow().Write(headerStyle).Write("Sub2", headerStyle).Write("Sub3", headerStyle)
                    .BeginRow().Write("Row3").Write(42).Write(-1)
                    .BeginRow().Write("Row4").SkipColumns(1).Write(1234)
                    .SkipRows(2)
                    .BeginRow().AddMergedCell(1, 2).Write("Row7").SkipColumns(1).Write(3.14159265359)
                    .BeginWorksheet("Sheet2")
                    .BeginRow().Write("Lorem ipsum dolor sit amet,")
                    .BeginRow().Write("consectetur adipiscing elit,")
                    .BeginRow().Write("sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            }
        }
    }
}
