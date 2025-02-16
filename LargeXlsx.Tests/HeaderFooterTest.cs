using NUnit.Framework;
using OfficeOpenXml;
using System.IO;
using Shouldly;

namespace LargeXlsx.Tests;

[TestFixture]
public static class HeaderFooterTest
{
    [Test]
    public static void Simple()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("HeaderFooterTest")
                .SetHeaderFooter(new XlsxHeaderFooter(
                    oddHeader: new XlsxHeaderFooterBuilder().Left().Text("LeftHeader").ToString(),
                    oddFooter: new XlsxHeaderFooterBuilder().Center().Text("CenterFooter").ToString()));
        }

        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.HeaderFooter.AlignWithMargins.ShouldBeTrue();
        sheet.HeaderFooter.ScaleWithDocument.ShouldBeTrue();
        sheet.HeaderFooter.differentFirst.ShouldBeFalse();
        sheet.HeaderFooter.differentOddEven.ShouldBeFalse();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.ShouldBe("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.ShouldBe("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.ShouldBeNull();
    }

    [Theory]
    public static void Flags(bool alignWithMargins, bool scaleWithDoc)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("HeaderFooterTest")
                .SetHeaderFooter(new XlsxHeaderFooter(
                    alignWithMargins: alignWithMargins,
                    scaleWithDoc: scaleWithDoc,
                    oddFooter: new XlsxHeaderFooterBuilder().Center().Text("Footer").ToString()));
        }

        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.HeaderFooter.AlignWithMargins.ShouldBe(alignWithMargins);
        sheet.HeaderFooter.ScaleWithDocument.ShouldBe(scaleWithDoc);
    }

    [Test]
    public static void DifferentFirst()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("HeaderFooterTest")
                .SetHeaderFooter(new XlsxHeaderFooter(
                    oddHeader: new XlsxHeaderFooterBuilder().Left().Text("LeftHeader").ToString(),
                    oddFooter: new XlsxHeaderFooterBuilder().Center().Text("CenterFooter").ToString(),
                    firstHeader: new XlsxHeaderFooterBuilder().Right().Text("FirstHeader").ToString()));
        }

        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.HeaderFooter.AlignWithMargins.ShouldBeTrue();
        sheet.HeaderFooter.ScaleWithDocument.ShouldBeTrue();
        sheet.HeaderFooter.differentFirst.ShouldBeTrue();
        sheet.HeaderFooter.differentOddEven.ShouldBeFalse();
        sheet.HeaderFooter.FirstHeader.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.FirstHeader.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.FirstHeader.RightAlignedText.ShouldBe("FirstHeader");
        sheet.HeaderFooter.FirstFooter.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.FirstFooter.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.FirstFooter.RightAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.ShouldBe("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.ShouldBe("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.ShouldBeNull();
    }

    [Test]
    public static void DifferentOddEven()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("HeaderFooterTest")
                .SetHeaderFooter(new XlsxHeaderFooter(
                    oddHeader: new XlsxHeaderFooterBuilder().Left().Text("LeftHeader").ToString(),
                    oddFooter: new XlsxHeaderFooterBuilder().Center().Text("CenterFooter").ToString(),
                    evenHeader: new XlsxHeaderFooterBuilder().Right().Text("EvenHeader").ToString()));
        }

        using var package = new ExcelPackage(stream);
        var sheet = package.Workbook.Worksheets[0];
        sheet.HeaderFooter.AlignWithMargins.ShouldBeTrue();
        sheet.HeaderFooter.ScaleWithDocument.ShouldBeTrue();
        sheet.HeaderFooter.differentFirst.ShouldBeFalse();
        sheet.HeaderFooter.differentOddEven.ShouldBeTrue();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.ShouldBe("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.ShouldBe("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.ShouldBeNull();
        sheet.HeaderFooter.EvenHeader.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.EvenHeader.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.EvenHeader.RightAlignedText.ShouldBe("EvenHeader");
        sheet.HeaderFooter.EvenFooter.LeftAlignedText.ShouldBeNull();
        sheet.HeaderFooter.EvenFooter.CenteredText.ShouldBeNull();
        sheet.HeaderFooter.EvenFooter.RightAlignedText.ShouldBeNull();
    }
}