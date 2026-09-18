/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2026 Salvatore ISAJA. All rights reserved.

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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NUnit.Framework;
using Shouldly;
using OpenXmlWorksheet = DocumentFormat.OpenXml.Spreadsheet.Worksheet;

namespace LargeXlsx.Tests;

[TestFixture]
public static class IgnoredErrorsTest
{
    [Test]
    public static void MissingIgnoredErrors()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().Write("123");

        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        GetWorksheet(spreadsheetDocument).Elements<IgnoredErrors>().ShouldBeEmpty();
    }

    [Test]
    public static void GroupsRangesWithTheSameErrorTypes()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter.BeginWorksheet("Sheet 1")
                .AddIgnoredErrors(2, 5, Limits.MaxRowCount - 1, 1, XlsxIgnoredError.NumberStoredAsText)
                .AddIgnoredErrors(2, 8, Limits.MaxRowCount - 1, 1, XlsxIgnoredError.NumberStoredAsText);
        }

        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var ignoredErrors = GetWorksheet(spreadsheetDocument)
            .Elements<IgnoredErrors>().Single()
            .Elements<IgnoredError>().ToArray();

        ignoredErrors.Length.ShouldBe(1);
        ignoredErrors[0].SequenceOfReferences!.InnerText.ShouldBe("E2:E1048576 H2:H1048576");
        ignoredErrors[0].NumberStoredAsText!.Value.ShouldBeTrue();
    }

    [Test]
    public static void GroupsSetsRegardlessOfOrderAndDuplicates()
    {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream))
        {
            writer.BeginWorksheet("Sheet 1").BeginRow().Write("123");
            writer.AddIgnoredErrors(2, 5, 99, 1,
                XlsxIgnoredError.Formula, XlsxIgnoredError.NumberStoredAsText);
            writer.AddIgnoredErrors(2, 8, 99, 1,
                XlsxIgnoredError.NumberStoredAsText, XlsxIgnoredError.Formula, XlsxIgnoredError.Formula);
            writer.AddIgnoredErrors(2, 5, 99, 1,
                XlsxIgnoredError.NumberStoredAsText, XlsxIgnoredError.Formula);
            writer.AddIgnoredErrors(2, 10, 99, 1, XlsxIgnoredError.NumberStoredAsText);
        }
        using var document = SpreadsheetDocument.Open(stream, false);
        var entries = GetWorksheet(document).Elements<IgnoredErrors>().Single().Elements<IgnoredError>().ToArray();
        entries.Length.ShouldBe(2);
        var combined = entries.Single(e => e.Formula?.Value == true);
        combined.SequenceOfReferences!.InnerText.ShouldBe("E2:E100 H2:H100");
        combined.NumberStoredAsText!.Value.ShouldBeTrue();
        var single = entries.Single(e => e.Formula == null);
        single.SequenceOfReferences!.InnerText.ShouldBe("J2:J100");
        single.NumberStoredAsText!.Value.ShouldBeTrue();
        new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document).ShouldBeEmpty();
    }

    [Test]
    public static void SupportsAllErrorTypesOnOneRange()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter.BeginWorksheet("Sheet 1")
                .AddIgnoredErrors(2, 3, 10, 4,
                    XlsxIgnoredError.CalculatedColumn,
                    XlsxIgnoredError.EmptyCellReference,
                    XlsxIgnoredError.EvaluationError,
                    XlsxIgnoredError.Formula,
                    XlsxIgnoredError.FormulaRange,
                    XlsxIgnoredError.ListDataValidation,
                    XlsxIgnoredError.NumberStoredAsText,
                    XlsxIgnoredError.TwoDigitTextYear,
                    XlsxIgnoredError.UnlockedFormula);
        }

        using var spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
        var ignoredError = GetWorksheet(spreadsheetDocument)
            .Elements<IgnoredErrors>().Single()
            .Elements<IgnoredError>().Single();

        ignoredError.SequenceOfReferences!.InnerText.ShouldBe("C2:F11");
        ignoredError.CalculatedColumn!.Value.ShouldBeTrue();
        ignoredError.EmptyCellReference!.Value.ShouldBeTrue();
        ignoredError.EvalError!.Value.ShouldBeTrue();
        ignoredError.Formula!.Value.ShouldBeTrue();
        ignoredError.FormulaRange!.Value.ShouldBeTrue();
        ignoredError.ListDataValidation!.Value.ShouldBeTrue();
        ignoredError.NumberStoredAsText!.Value.ShouldBeTrue();
        ignoredError.TwoDigitTextYear!.Value.ShouldBeTrue();
        ignoredError.UnlockedFormula!.Value.ShouldBeTrue();
    }

    [TestCase(0, 1, 1, 1)]
    [TestCase(1, 0, 1, 1)]
    [TestCase(1, 1, 0, 1)]
    [TestCase(1, 1, 1, 0)]
    [TestCase(Limits.MaxRowCount, 1, 2, 1)]
    [TestCase(1, Limits.MaxColumnCount, 1, 2)]
    public static void RejectsInvalidRanges(int fromRow, int fromColumn, int rowCount, int columnCount)
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1");

        Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddIgnoredErrors(
            fromRow, fromColumn, rowCount, columnCount, XlsxIgnoredError.NumberStoredAsText));
    }

    [Test]
    public static void RequiresAtLeastOneErrorType()
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1");

        Should.Throw<ArgumentException>(() => xlsxWriter.AddIgnoredErrors(1, 1, 1, 1));
    }

    [TestCase(-1)]
    [TestCase(512)]
    [TestCase(512 | (int)XlsxIgnoredError.NumberStoredAsText)]
    [TestCase(int.MaxValue)]
    public static void RejectsUnknownErrorType(int value)
    {
        using var stream = new MemoryStream();
        using var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1");

        Should.Throw<ArgumentOutOfRangeException>(() => xlsxWriter.AddIgnoredErrors(
            1, 1, 1, 1, (XlsxIgnoredError)value));
    }


    [Test]
    public static void CopiesErrorTypesAndRemovesDuplicates()
    {
        using var stream = new MemoryStream();
        var errors = new[] { XlsxIgnoredError.NumberStoredAsText, XlsxIgnoredError.NumberStoredAsText };
        using (var writer = new XlsxWriter(stream))
        {
            writer.BeginWorksheet("Sheet 1").AddIgnoredErrors(1, 1, 1, 1, errors);
            errors[0] = XlsxIgnoredError.Formula;
        }
        using var document = SpreadsheetDocument.Open(stream, false);
        var entry = GetWorksheet(document).Elements<IgnoredErrors>().Single().Elements<IgnoredError>().Single();
        entry.NumberStoredAsText!.Value.ShouldBeTrue();
        entry.Formula.ShouldBeNull();
        entry.GetAttributes().Count.ShouldBe(2);
    }

    [Test]
    public static void CurrentPositionOverloadsAndWorksheetsAreIndependent()
    {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream))
        {
            writer.BeginWorksheet("Sheet 1").BeginRow().Write("123");
            writer.AddIgnoredErrors(XlsxIgnoredError.NumberStoredAsText);
            writer.AddIgnoredErrors(2, 3, XlsxIgnoredError.Formula, XlsxIgnoredError.TwoDigitTextYear);
            writer.BeginWorksheet("Sheet 2").BeginRow().Write("456");
        }
        using var document = SpreadsheetDocument.Open(stream, false);
        var entries = GetWorksheet(document).Elements<IgnoredErrors>().Single().Elements<IgnoredError>().ToArray();
        entries[0].SequenceOfReferences!.InnerText.ShouldBe("B1");
        entries[0].NumberStoredAsText!.Value.ShouldBeTrue();
        entries[1].SequenceOfReferences!.InnerText.ShouldBe("B1:D2");
        entries[1].Formula!.Value.ShouldBeTrue();
        entries[1].TwoDigitTextYear!.Value.ShouldBeTrue();
        entries[1].NumberStoredAsText.ShouldBeNull();
        document.WorkbookPart!.WorksheetParts.Count(p => p.Worksheet!.Elements<IgnoredErrors>().Any()).ShouldBe(1);
        new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document).ShouldBeEmpty();
    }

    [Test]
    public static void RejectsNullErrorTypes()
    {
        using var stream = new MemoryStream();
        using var writer = new XlsxWriter(stream);
        writer.BeginWorksheet("Sheet 1");
        Should.Throw<ArgumentNullException>(() => writer.AddIgnoredErrors(1, 1, 1, 1, null!));
    }

    [Test]
    public static void CombinedFlagsAndSeparateArgumentsShareAGroup()
    {
        using var stream = new MemoryStream();
        using (var writer = new XlsxWriter(stream))
        {
            writer.BeginWorksheet("Sheet 1").BeginRow().Write("123");
            writer.AddIgnoredErrors(2, 5, 99, 1,
                XlsxIgnoredError.Formula | XlsxIgnoredError.NumberStoredAsText);
            writer.AddIgnoredErrors(2, 8, 99, 1,
                XlsxIgnoredError.NumberStoredAsText, XlsxIgnoredError.Formula);
        }
        using var document = SpreadsheetDocument.Open(stream, false);
        var entry = GetWorksheet(document).Elements<IgnoredErrors>().Single().Elements<IgnoredError>().Single();
        entry.SequenceOfReferences!.InnerText.ShouldBe("E2:E100 H2:H100");
        entry.Formula!.Value.ShouldBeTrue();
        entry.NumberStoredAsText!.Value.ShouldBeTrue();
        new DocumentFormat.OpenXml.Validation.OpenXmlValidator().Validate(document).ShouldBeEmpty();
    }

    [Test]
    public static void RejectsZeroErrorFlags()
    {
        using var stream = new MemoryStream();
        using var writer = new XlsxWriter(stream);
        writer.BeginWorksheet("Sheet 1");
        Should.Throw<ArgumentException>(() => writer.AddIgnoredErrors(1, 1, 1, 1, (XlsxIgnoredError)0));
    }

    private static OpenXmlWorksheet GetWorksheet(SpreadsheetDocument spreadsheetDocument)
    {
        var sheetId = spreadsheetDocument.WorkbookPart!.Workbook!.Sheets!
            .Elements<Sheet>().Single(s => s.Name == "Sheet 1").Id!.ToString()!;
        return ((WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheetId)).Worksheet!;
    }
}
