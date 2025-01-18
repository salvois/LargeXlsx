/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025 Salvatore ISAJA. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice,
this list of conditions and the following disclaimer in the documentation
and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED THE COPYRIGHT HOLDER ``AS IS'' AND ANY EXPRESS
OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES
OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN
NO EVENT SHALL THE COPYRIGHT HOLDER BE LIABLE FOR ANY DIRECT,
INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
*/
using System.IO;
using LargeXlsx;

namespace Examples;

public static class NumberFormats
{
    public static void Run()
    {
        using var stream = new FileStream($"{nameof(NumberFormats)}.xlsx", FileMode.Create, FileAccess.Write);
        using var xlsxWriter = new XlsxWriter(stream);
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
        var textStyle = XlsxStyle.Default.With(XlsxNumberFormat.Text);
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
            .BeginRow().Write(nameof(XlsxNumberFormat.Text)).Write(1234.5678, textStyle).Write(1.2, textStyle)
            .BeginRow().Write(customNumberFormat1.FormatCode).Write(1234.5678, customStyle1).Write(1.2, customStyle1)
            .BeginRow().Write(customNumberFormat2.FormatCode).Write(1234.5678, customStyle2).Write(1.2, customStyle2);
    }
}