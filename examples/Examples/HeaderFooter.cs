using System.IO;
using LargeXlsx;

namespace Examples;

public static class HeaderFooter
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(HeaderFooter)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);

        var headerFooter = new XlsxHeaderFooter(
            oddHeader: new XlsxHeaderFooterBuilder()
                .Left().Bold().Text(nameof(HeaderFooter)).Bold().Text(" example")
                .ToString(),
            oddFooter: new XlsxHeaderFooterBuilder()
                .Left().Text("Page ").PageNumber(-42).Text(" of ").NumberOfPages()
                .Center().SheetName()
                .Right().FileName()
                .ToString());

        xlsxWriter.BeginWorksheet("Sheet1")
            .SetHeaderFooter(headerFooter)
            .BeginRow().Write("A1").Write("B1").BeginRow().Write("A2").Write("B2");
    }
}