/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2023 Salvatore ISAJA. All rights reserved.

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

using System;
using System.IO;
using System.Linq;
using System.Xml;
using FluentAssertions;
using NUnit.Framework;
using OfficeOpenXml;

namespace LargeXlsx.Tests;

[TestFixture]
public static class IgnoreErrorsTest
{
    [Test]
    public static void MissingIgnoreErrors()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("test")
                .BeginRow()
                .Write();
        }

        using (var package = new ExcelPackage(stream))
        {
            var xmlDocument = package.Workbook.Worksheets[0].WorksheetXml;

            var worksheets = xmlDocument.GetElementsByTagName("worksheet");

            worksheets.Count.Should().Be(1);

            var worksheet = worksheets.Item(0)!;

            var ignoredErrorsElement = worksheet.ChildNodes.Cast<XmlNode>().FirstOrDefault(z => string.Equals(z.Name, "ignoredErrors", StringComparison.InvariantCultureIgnoreCase));

            ignoredErrorsElement.Should().BeNull();
        }
    }

    [Test]
    public static void IgnoreErrors()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter
                .BeginWorksheet("test")
                .BeginRow()
                .Write()
                .BeginRow()
                .Write()
                .Write()
                .AddIgnoreError(1, 1, 1, 1, XlsxDataIgnoreError.ErrorType.NumberStoredAsText)
                .AddIgnoreError(2, 1, 1, 2, XlsxDataIgnoreError.ErrorType.NumberStoredAsText);
        }

        using (var package = new ExcelPackage(stream))
        {
            var xmlDocument = package.Workbook.Worksheets[0].WorksheetXml;
            
            var worksheets = xmlDocument.GetElementsByTagName("worksheet");

            worksheets.Count.Should().Be(1);

            var worksheet = worksheets.Item(0)!;
            
            var ignoredErrorsElement = worksheet.ChildNodes.Cast<XmlNode>().FirstOrDefault(z => string.Equals(z.Name, "ignoredErrors", StringComparison.InvariantCultureIgnoreCase));

            ignoredErrorsElement.Should().NotBeNull();

            foreach (var element in ignoredErrorsElement!.ChildNodes.Cast<XmlNode>())
            {
                element.Name.ToLower().Should().Be("ignoredError".ToLower());
                element.Attributes.Should().NotBeNull();
                
                var numberStoredAsTextAttr = element.Attributes!.Cast<XmlAttribute>().FirstOrDefault(z => string.Equals(z.Name, "numberStoredAsText", StringComparison.InvariantCultureIgnoreCase));
                numberStoredAsTextAttr.Should().NotBeNull();
                numberStoredAsTextAttr!.Value.Should().Be("1");
                
                var sqrefAttr = element.Attributes!.Cast<XmlAttribute>().FirstOrDefault(z => string.Equals(z.Name, "sqref", StringComparison.InvariantCultureIgnoreCase));
                sqrefAttr.Should().NotBeNull();
            }


            var actualAddresses = ignoredErrorsElement!.ChildNodes.Cast<XmlNode>().SelectMany(z => z.Attributes!.Cast<XmlAttribute>())
                .Where(z => string.Equals(z.Name, "sqref", StringComparison.InvariantCultureIgnoreCase))
                .Select(z => z.Value)
                .OrderBy(z=>z)
                .ToList();

            var expectedAddresses = new[] {"A1", "A2:B2"}.OrderBy(z => z).ToList();

            actualAddresses.Should().BeEquivalentTo(expectedAddresses);
        }
    }
}