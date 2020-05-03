using System;
using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class Columns
    {
        public static void Run()
        {
            var rnd = new Random();
            using (var stream = new FileStream($"{nameof(Columns)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var blueStyle = new XlsxStyle(
                    new XlsxFont(XlsxFont.Default.FontName, XlsxFont.Default.FontSize, "ffffff"),
                    new XlsxFill(XlsxFill.Pattern.Solid, "004586"),
                    XlsxBorder.None,
                    XlsxNumberFormat.General);

                xlsxWriter
                    .BeginWorksheet("Sheet 1", columns: new[]
                    {
                        XlsxColumn.Formatted(count: 2, width: 20),
                        XlsxColumn.Unformatted(3),
                        XlsxColumn.Formatted(style: blueStyle, width: 9),
                        XlsxColumn.Formatted(hidden: true, width: 0)
                    });
                for (var i = 0; i < 10; i++)
                {
                    xlsxWriter.BeginRow();
                    for (var j = 0; j < 10; j++)
                        xlsxWriter.Write(rnd.Next());
                }
            }
        }
    }
}