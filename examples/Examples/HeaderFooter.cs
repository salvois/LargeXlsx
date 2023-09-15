using System;
using System.Drawing;
using System.IO;
using LargeXlsx;

namespace Examples;

public static class HeaderFooter
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(HeaderFooter)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
        var headerStyle = new XlsxStyle(
            new XlsxFont("Segoe UI", 9, Color.White, bold: true),
            new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
            XlsxStyle.Default.Border,
            XlsxStyle.Default.NumberFormat,
            XlsxAlignment.Default);
        var highlightStyle = XlsxStyle.Default.With(new XlsxFill(Color.FromArgb(0xff, 0xff, 0x88)));
        var dateStyle = XlsxStyle.Default.With(XlsxNumberFormat.ShortDateTime);
        var borderedStyle = highlightStyle.With(XlsxBorder.Around(new XlsxBorder.Line(Color.DeepPink, XlsxBorder.Style.Dashed)));
        var hyperlinkStyle = XlsxStyle.Default.With(XlsxFont.Default.WithUnderline().With(Color.Blue));

        var headerFooter = 
            new XlsxHeaderFooter(
                new XlsxHeaderFooterText($"{XlsxHeaderFooter.Bold}{nameof(HeaderFooter)} Example"),
                new XlsxHeaderFooterText(
                    $"Page {XlsxHeaderFooter.PageNumber} of {XlsxHeaderFooter.NumberOfPages}",
                    XlsxHeaderFooter.SheetName,
                    XlsxHeaderFooter.FileName));
        
        xlsxWriter
            .BeginWorksheet("Sheet 1", columns: new[] { XlsxColumn.Unformatted(count: 2), XlsxColumn.Formatted(width: 20) })
            .SetHeaderFooter(headerFooter)
            .SetDefaultStyle(headerStyle)
            .BeginRow().AddMergedCell(2, 1).Write("Col1").Write("Top2").Write("Top3")
            .BeginRow().Write().Write("Col2").Write("Col3")
            .SetDefaultStyle(XlsxStyle.Default)
            .BeginRow().Write("Row3").Write(42).WriteFormula(
                $"{xlsxWriter.GetRelativeColumnName(-1)}{xlsxWriter.CurrentRowNumber}*10", highlightStyle)
            .BeginRow().Write("Row4").SkipColumns(1).Write(new DateTime(2020, 5, 6, 18, 27, 0), dateStyle)
            .SkipRows(2)
            .BeginRow().Write("Row7", borderedStyle, columnSpan: 2).Write(3.14159265359)
            .BeginRow().Write("Bold").Write().Write("Be bold", XlsxStyle.Default.With(XlsxFont.Default.WithBold()))
            .BeginRow().Write("Italic").Write().Write("Be italic", XlsxStyle.Default.With(XlsxFont.Default.WithItalic()))
            .BeginRow().Write("Strike").Write().Write("Be struck", XlsxStyle.Default.With(XlsxFont.Default.WithStrike()))
            .BeginRow().Write("Underline").Write().Write("Single", XlsxStyle.Default.With(XlsxFont.Default.WithUnderline()))
            .BeginRow().Write("Underline").Write().Write("Double", XlsxStyle.Default.With(XlsxFont.Default.WithUnderline(XlsxFont.Underline.Double)))
            .BeginRow().Write("Underline").Write().Write("SingleAccounting", XlsxStyle.Default.With(XlsxFont.Default.WithUnderline(XlsxFont.Underline.SingleAccounting)))
            .BeginRow().Write("Underline").Write().Write("DoubleAccounting", XlsxStyle.Default.With(XlsxFont.Default.WithUnderline(XlsxFont.Underline.DoubleAccounting)))
            .BeginRow().Write("Hyperlink").Write().WriteFormula("HYPERLINK(\"https://github.com/salvois/LargeXlsx\")", hyperlinkStyle)
            .BeginRow().Write("Hyperlink w/alias").Write().WriteFormula("HYPERLINK(\"https://github.com/salvois/LargeXlsx\", \"LargeXlsx on GitHub\")", hyperlinkStyle)
            .BeginRow().Write("Boolean").Write(false).Write(true)
            .SetAutoFilter(2, 1, xlsxWriter.CurrentRowNumber - 1, 3);
    }
}