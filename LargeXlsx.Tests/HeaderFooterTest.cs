using NUnit.Framework;
using OfficeOpenXml;
using System.IO;
using FluentAssertions;

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
        sheet.HeaderFooter.AlignWithMargins.Should().BeTrue();
        sheet.HeaderFooter.ScaleWithDocument.Should().BeTrue();
        sheet.HeaderFooter.differentFirst.Should().BeFalse();
        sheet.HeaderFooter.differentOddEven.Should().BeFalse();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.Should().Be("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.Should().BeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.Should().Be("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.Should().BeNull();
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
        sheet.HeaderFooter.AlignWithMargins.Should().Be(alignWithMargins);
        sheet.HeaderFooter.ScaleWithDocument.Should().Be(scaleWithDoc);
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
        sheet.HeaderFooter.AlignWithMargins.Should().BeTrue();
        sheet.HeaderFooter.ScaleWithDocument.Should().BeTrue();
        sheet.HeaderFooter.differentFirst.Should().BeTrue();
        sheet.HeaderFooter.differentOddEven.Should().BeFalse();
        sheet.HeaderFooter.FirstHeader.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.FirstHeader.CenteredText.Should().BeNull();
        sheet.HeaderFooter.FirstHeader.RightAlignedText.Should().Be("FirstHeader");
        sheet.HeaderFooter.FirstFooter.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.FirstFooter.CenteredText.Should().BeNull();
        sheet.HeaderFooter.FirstFooter.RightAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.Should().Be("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.Should().BeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.Should().Be("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.Should().BeNull();
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
        sheet.HeaderFooter.AlignWithMargins.Should().BeTrue();
        sheet.HeaderFooter.ScaleWithDocument.Should().BeTrue();
        sheet.HeaderFooter.differentFirst.Should().BeFalse();
        sheet.HeaderFooter.differentOddEven.Should().BeTrue();
        sheet.HeaderFooter.OddHeader.LeftAlignedText.Should().Be("LeftHeader");
        sheet.HeaderFooter.OddHeader.CenteredText.Should().BeNull();
        sheet.HeaderFooter.OddHeader.RightAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.OddFooter.CenteredText.Should().Be("CenterFooter");
        sheet.HeaderFooter.OddFooter.RightAlignedText.Should().BeNull();
        sheet.HeaderFooter.EvenHeader.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.EvenHeader.CenteredText.Should().BeNull();
        sheet.HeaderFooter.EvenHeader.RightAlignedText.Should().Be("EvenHeader");
        sheet.HeaderFooter.EvenFooter.LeftAlignedText.Should().BeNull();
        sheet.HeaderFooter.EvenFooter.CenteredText.Should().BeNull();
        sheet.HeaderFooter.EvenFooter.RightAlignedText.Should().BeNull();
    }
}