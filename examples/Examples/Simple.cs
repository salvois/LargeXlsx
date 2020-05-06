using System;
using System.Drawing;
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
                var whiteFont = new XlsxFont("Segoe UI", 9, Color.White, bold: true);
                var blueFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0, 0x45, 0x86));
                var yellowFill = new XlsxFill(XlsxFill.Pattern.Solid, Color.FromArgb(0xff, 0xff, 0x88));
                var headerStyle = new XlsxStyle(whiteFont, blueFill, XlsxBorder.None, XlsxNumberFormat.General);
                var highlightStyle = new XlsxStyle(XlsxFont.Default, yellowFill, XlsxBorder.None, XlsxNumberFormat.General);
                var dateStyle = new XlsxStyle(XlsxStyle.Default.Font, XlsxStyle.Default.Fill, XlsxStyle.Default.Border, XlsxNumberFormat.ShortDateTime);

                xlsxWriter
                    .BeginWorksheet("Sheet&'<1>\"", columns: new [] { XlsxColumn.Unformatted(count: 2), XlsxColumn.Formatted(width: 20) })
                    .SetDefaultStyle(headerStyle)
                    .BeginRow().Write("Col<1>").Write("Col2").Write("Col&3")
                    .BeginRow().Write().Write("Sub2").Write("Sub3")
                    .SetDefaultStyle(XlsxStyle.Default)
                    .BeginRow().Write("Row3").Write(42).Write(-1, highlightStyle)
                    .BeginRow().Write("Row4").SkipColumns(1).Write(new DateTime(2020, 5, 6, 18, 27, 0), dateStyle)
                    .SkipRows(2)
                    .BeginRow().Write("Row7", columnSpan: 2).Write(3.14159265359)
                    .SetAutoFilter(1, 1, xlsxWriter.CurrentRowNumber, 3)
                    .BeginWorksheet("Sheet2")
                    .BeginRow().Write("Lorem ipsum dolor sit amet,")
                    .BeginRow().Write("consectetur adipiscing elit,")
                    .BeginRow().Write("sed do eiusmod tempor incididunt ut labore et dolore magna aliqua.");
            }
        }
    }
}
