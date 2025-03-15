using System.IO;
using LargeXlsx;

namespace Examples;

public static class HeaderFooterPageBreaks
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(HeaderFooterPageBreaks)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);

        var headerFooter = new XlsxHeaderFooter(
            oddHeader: new XlsxHeaderFooterBuilder()
                .Left().Bold().Text(nameof(HeaderFooterPageBreaks)).Bold().Text(" example")
                .ToString(),
            oddFooter: new XlsxHeaderFooterBuilder()
                .Left().Text("Page ").PageNumber(-42).Text(" of ").NumberOfPages()
                .Center().SheetName()
                .Right().FileName()
                .ToString());

        xlsxWriter.BeginWorksheet("Sheet1")
            .SetHeaderFooter(headerFooter)
            .BeginRow().Write("A1").Write("B1").AddColumnPageBreak().Write("C1")
            .BeginRow().Write("A2").Write("B2").Write("C2")
            .BeginRow().Write("A3").AddRowPageBreak().Write("B3").Write("C3")
            .BeginRow().Write("A4").Write("B4").Write("C4");
    }
}