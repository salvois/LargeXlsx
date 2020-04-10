using System.IO;
using LargeXlsx;

namespace Examples
{
    public static class NumberFormats
    {
        public static void Run()
        {
            using (var stream = new FileStream($"{nameof(NumberFormats)}.xlsx", FileMode.Create, FileAccess.Write))
            using (var xlsxWriter = new XlsxWriter(stream))
            {
                var customNumberFormat1 = new XlsxNumberFormat("0.0%");
                var customNumberFormat2 = new XlsxNumberFormat("#,##0.00####");

                var generalStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.General);
                var twoDecimalStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.TwoDecimal);
                var thousandTwoDecimalStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.ThousandTwoDecimal);
                var percentageStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.Percentage);
                var scientificStyle = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, XlsxNumberFormat.Scientific);
                var customStyle1 = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, customNumberFormat1);
                var customStyle2 = new XlsxStyle(XlsxFont.Default, XlsxFill.None, XlsxBorder.None, customNumberFormat2);

                xlsxWriter
                    .BeginWorksheet("Sheet1")
                    .BeginRow().Write(nameof(XlsxNumberFormat.General)).Write(1234.5678, generalStyle).Write(1.2, generalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.TwoDecimal)).Write(1234.5678, twoDecimalStyle).Write(1.2, twoDecimalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.ThousandTwoDecimal)).Write(1234.5678, thousandTwoDecimalStyle).Write(1.2, thousandTwoDecimalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.Percentage)).Write(1234.5678, percentageStyle).Write(1.2, percentageStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.Scientific)).Write(1234.5678, scientificStyle).Write(1.2, scientificStyle)
                    .BeginRow().Write(customNumberFormat1.FormatCode).Write(1234.5678, customStyle1).Write(1.2, customStyle1)
                    .BeginRow().Write(customNumberFormat2.FormatCode).Write(1234.5678, customStyle2).Write(1.2, customStyle2);
            }
        }
    }
}