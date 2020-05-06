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

                var generalStyle = XlsxStyle.Default.With(XlsxNumberFormat.General);
                var integerStyle = XlsxStyle.Default.With(XlsxNumberFormat.Integer);
                var twoDecimalStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimal);
                var thousandIntegerStyle = XlsxStyle.Default.With(XlsxNumberFormat.ThousandInteger);
                var thousandTwoDecimalStyle = XlsxStyle.Default.With(XlsxNumberFormat.ThousandTwoDecimal);
                var integerPercentageStyle = XlsxStyle.Default.With(XlsxNumberFormat.IntegerPercentage);
                var twoDecimalPercentageStyle = XlsxStyle.Default.With(XlsxNumberFormat.TwoDecimalPercentage);
                var scientificStyle = XlsxStyle.Default.With(XlsxNumberFormat.Scientific);
                var shortDateStyle = XlsxStyle.Default.With(XlsxNumberFormat.ShortDate);
                var shortDateTimeStyle = XlsxStyle.Default.With(XlsxNumberFormat.ShortDateTime);
                var customStyle1 = XlsxStyle.Default.With(customNumberFormat1);
                var customStyle2 = XlsxStyle.Default.With(customNumberFormat2);

                xlsxWriter
                    .BeginWorksheet("Sheet1")
                    .BeginRow().Write(nameof(XlsxNumberFormat.General)).Write(1234.5678, generalStyle).Write(1.2, generalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.Integer)).Write(1234.5678, integerStyle).Write(1.2, integerStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.TwoDecimal)).Write(1234.5678, twoDecimalStyle).Write(1.2, twoDecimalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.ThousandInteger)).Write(1234.5678, thousandIntegerStyle).Write(1.2, thousandIntegerStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.ThousandTwoDecimal)).Write(1234.5678, thousandTwoDecimalStyle).Write(1.2, thousandTwoDecimalStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.IntegerPercentage)).Write(1234.5678, integerPercentageStyle).Write(1.2, integerPercentageStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.TwoDecimalPercentage)).Write(1234.5678, twoDecimalPercentageStyle).Write(1.2, twoDecimalPercentageStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.Scientific)).Write(1234.5678, scientificStyle).Write(1.2, scientificStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.ShortDate)).Write(1234.5678, shortDateStyle).Write(1.2, shortDateStyle)
                    .BeginRow().Write(nameof(XlsxNumberFormat.ShortDateTime)).Write(1234.5678, shortDateTimeStyle).Write(1.2, shortDateTimeStyle)
                    .BeginRow().Write(customNumberFormat1.FormatCode).Write(1234.5678, customStyle1).Write(1.2, customStyle1)
                    .BeginRow().Write(customNumberFormat2.FormatCode).Write(1234.5678, customStyle2).Write(1.2, customStyle2);
            }
        }

        private static XlsxStyle With(this XlsxStyle style, XlsxNumberFormat numberFormat) =>
            new XlsxStyle(style.Font, style.Fill, style.Border, numberFormat, style.Alignment);
    }
}