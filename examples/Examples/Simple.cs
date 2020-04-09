using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class Simple
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(Simple)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var whiteFont = xlsxWriter.Stylesheet.CreateFont("Segoe UI", 9, "ffffff", bold: true);
                var blueFill = xlsxWriter.Stylesheet.CreateSolidFill("004586");
                var yellowFill = xlsxWriter.Stylesheet.CreateSolidFill("ffff88");
                var headerStyle = xlsxWriter.Stylesheet.CreateStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                var highlightStyle = xlsxWriter.Stylesheet.CreateStyle(XlsxFont.Default, yellowFill, XlsxBorder.None, XlsxNumberFormat.General);

                xlsxWriter
                    .BeginWorksheet("Sheet&'<1>\"")
                    .SetDefaultStyle(headerStyle)
                    .BeginRow().Write("Col<1>").Write("Col2").Write("Col&3")
                    .BeginRow().Write().Write("Sub2").Write("Sub3")
                    .SetDefaultStyle(XlsxStyle.Default)
                    .BeginRow().Write("Row3").Write(42).Write(-1, highlightStyle)
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
