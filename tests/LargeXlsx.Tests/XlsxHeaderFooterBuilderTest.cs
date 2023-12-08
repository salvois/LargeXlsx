using FluentAssertions;
using NUnit.Framework;

namespace LargeXlsx.Tests;

[TestFixture]
public static class XlsxHeaderFooterBuilderTest
{
    [Test]
    public static void Simple() =>
        new XlsxHeaderFooterBuilder()
            .Left().Bold().Italic().Underline().Text("Left&Formatted")
            .Center().DoubleUnderline().Subscript().PageNumber().NumberOfPages().CurrentDate().CurrentTime().FilePath().FileName()
            .Right().Superscript().StrikeThrough().SheetName()
            .ToString()
            .Should().Be("&L&B&I&ULeft&&Formatted"
                         + "&C&E&Y&P&N&D&T&Z&F"
                         + "&R&X&S&A");

    [TestCase(0, "&P")]
    [TestCase(1, "&P+1")]
    [TestCase(-1, "&P-1")]
    [TestCase(42, "&P+42")]
    [TestCase(-69, "&P-69")]
    public static void PageNumber(int offset, string expected) => 
        new XlsxHeaderFooterBuilder().PageNumber(offset).ToString().Should().Be(expected);

    [TestCase(1, "&1")]
    [TestCase(42, "&42")]
    public static void FontSize(int points, string expected) => 
        new XlsxHeaderFooterBuilder().FontSize(points).ToString().Should().Be(expected);

    [TestCase("Times New Roman", false, false, "&\"Times New Roman,Regular\"")]
    [TestCase("Times New Roman", true, false, "&\"Times New Roman,Bold\"")]
    [TestCase("Times New Roman", false, true, "&\"Times New Roman,Italic\"")]
    [TestCase("Times New Roman", true, true, "&\"Times New Roman,Bold Italic\"")]
    public static void Font(string name, bool bold, bool italic, string expected) => 
        new XlsxHeaderFooterBuilder().Font(name, bold, italic).ToString().Should().Be(expected);

    [TestCase(false, false, "&\"-,Regular\"")]
    [TestCase(true, false, "&\"-,Bold\"")]
    [TestCase(false, true, "&\"-,Italic\"")]
    [TestCase(true, true, "&\"-,Bold Italic\"")]
    public static void FontWithoutName(bool bold, bool italic, string expected) => 
        new XlsxHeaderFooterBuilder().Font(bold, italic).ToString().Should().Be(expected);
}