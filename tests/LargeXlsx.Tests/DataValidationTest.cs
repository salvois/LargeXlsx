/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2022 Salvatore ISAJA. All rights reserved.

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
using System.Drawing;
using System.IO;
using FluentAssertions;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.DataValidation;

namespace LargeXlsx.Tests;

[TestFixture]
public static class DataValidationTest
{
    [Test]
    public static void ListAsFormula()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var dataValidation = new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, formula1: "\"Lorem,Ipsum,Dolor\"");
            xlsxWriter.SetDefaultStyle(XlsxStyle.Default.With(new XlsxFill(Color.OldLace))).BeginWorksheet("Sheet 1")
                .BeginRow().AddDataValidation(dataValidation).Write().SkipColumns(1).AddDataValidation(1, 2, dataValidation).Write(repeatCount: 2);
        }
        using (var package = new ExcelPackage(stream))
        {
            var dataValidation = package.Workbook.Worksheets[0].DataValidations[0] as ExcelDataValidationList;
            dataValidation.Should().NotBeNull();
            dataValidation!.ValidationType.Should().Be(ExcelDataValidationType.List);
            dataValidation.Address.Address.Should().Be("A1 C1:D1");
            dataValidation.Formula.Values.Should().BeEquivalentTo("Lorem", "Ipsum", "Dolor");
        }
    }

    [Test]
    public static void ListAsChoices()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(XlsxDataValidation.List(new[] { "Lorem", "Ipsum", "Dolor" })).Write();
        using var package = new ExcelPackage(stream);
        var dataValidation = package.Workbook.Worksheets[0].DataValidations[0] as ExcelDataValidationList;
        dataValidation.Should().NotBeNull();
        dataValidation!.ValidationType.Should().Be(ExcelDataValidationType.List);
        dataValidation.Address.Address.Should().Be("A1");
        dataValidation.Formula.Values.Should().BeEquivalentTo("Lorem", "Ipsum", "Dolor");
    }

    [Test]
    public static void ListByRange()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            var dataValidation = new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, formula1: "=Choices!A1:A3");
            xlsxWriter
                .BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(dataValidation).Write(XlsxStyle.Default.With(new XlsxFill(Color.OldLace)))
                .BeginWorksheet("Choices").BeginRow().Write(3.141592).BeginRow().Write("Lorem").BeginRow().Write("Ipsum, Dolor");
        }
        using (var package = new ExcelPackage(stream))
        {
            var dataValidation = package.Workbook.Worksheets[0].DataValidations[0] as ExcelDataValidationList;
            dataValidation.Should().NotBeNull();
            dataValidation!.ValidationType.Should().Be(ExcelDataValidationType.List);
            dataValidation.Address.Address.Should().Be("A1");
            dataValidation.Formula.ExcelFormula.Should().Be("=Choices!A1:A3");
        }
    }

    [Test]
    public static void TwoFormulas()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(
                new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.Decimal, formula1: "1", formula2: "10"));
        using var package = new ExcelPackage(stream);
        var dataValidation = package.Workbook.Worksheets[0].DataValidations[0] as ExcelDataValidationDecimal;
        dataValidation.Should().NotBeNull();
        dataValidation!.Formula.Value.Should().Be(1.0);
        dataValidation.Formula2.Value.Should().Be(10.0);
    }

    [Test]
    public static void Messages()
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
        {
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(
                new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List,
                    showErrorMessage: true, errorTitle: "Error title", error: "A very informative error message",
                    showInputMessage: true, promptTitle: "Prompt title", prompt: "A very enlightening prompt"));
        }
        using var package = new ExcelPackage(stream);
        var dataValidation = (ExcelDataValidationList)package.Workbook.Worksheets[0].DataValidations[0];
        dataValidation.ShowErrorMessage.Should().BeTrue();
        dataValidation.ErrorTitle.Should().Be("Error title");
        dataValidation.Error.Should().Be("A very informative error message");
        dataValidation.ShowInputMessage.Should().BeTrue();
        dataValidation.PromptTitle.Should().Be("Prompt title");
        dataValidation.Prompt.Should().Be("A very enlightening prompt");
    }

    [TestCase(XlsxDataValidation.ErrorStyle.Information, ExcelDataValidationWarningStyle.information)]
    [TestCase(XlsxDataValidation.ErrorStyle.Warning, ExcelDataValidationWarningStyle.warning)]
    [TestCase(XlsxDataValidation.ErrorStyle.Stop, ExcelDataValidationWarningStyle.stop)]
    public static void ErrorStyles(XlsxDataValidation.ErrorStyle errorStyle, ExcelDataValidationWarningStyle expected)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.List, errorStyle: errorStyle));
        using var package = new ExcelPackage(stream);
        var dataValidation = (ExcelDataValidationList)package.Workbook.Worksheets[0].DataValidations[0];
        dataValidation.ErrorStyle.Should().Be(expected);
    }

    [TestCase(XlsxDataValidation.ValidationType.Custom, eDataValidationType.Custom)]
    [TestCase(XlsxDataValidation.ValidationType.Date, eDataValidationType.DateTime)]
    [TestCase(XlsxDataValidation.ValidationType.Decimal, eDataValidationType.Decimal)]
    [TestCase(XlsxDataValidation.ValidationType.List, eDataValidationType.List)]
    [TestCase(XlsxDataValidation.ValidationType.TextLength, eDataValidationType.TextLength)]
    [TestCase(XlsxDataValidation.ValidationType.Time, eDataValidationType.Time)]
    [TestCase(XlsxDataValidation.ValidationType.Whole, eDataValidationType.Whole)]
    // XlsxDataValidation.ValidationType.None has no corresponding representation in EPPlus
    public static void ValidationType(XlsxDataValidation.ValidationType validationType, eDataValidationType expected)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(new XlsxDataValidation(validationType: validationType));
        using (var package = new ExcelPackage(stream))
            package.Workbook.Worksheets[0].DataValidations[0].ValidationType.Type.Should().Be(expected);
    }

    [TestCase(XlsxDataValidation.Operator.Between, ExcelDataValidationOperator.between)]
    [TestCase(XlsxDataValidation.Operator.Equal, ExcelDataValidationOperator.equal)]
    [TestCase(XlsxDataValidation.Operator.GreaterThan, ExcelDataValidationOperator.greaterThan)]
    [TestCase(XlsxDataValidation.Operator.GreaterThanOrEqual, ExcelDataValidationOperator.greaterThanOrEqual)]
    [TestCase(XlsxDataValidation.Operator.LessThan, ExcelDataValidationOperator.lessThan)]
    [TestCase(XlsxDataValidation.Operator.LessThanOrEqual, ExcelDataValidationOperator.lessThanOrEqual)]
    [TestCase(XlsxDataValidation.Operator.NotBetween, ExcelDataValidationOperator.notBetween)]
    [TestCase(XlsxDataValidation.Operator.NotEqual, ExcelDataValidationOperator.notEqual)]
    public static void Operator(XlsxDataValidation.Operator operatorType, ExcelDataValidationOperator expected)
    {
        using var stream = new MemoryStream();
        using (var xlsxWriter = new XlsxWriter(stream))
            xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().AddDataValidation(new XlsxDataValidation(validationType: XlsxDataValidation.ValidationType.Decimal, operatorType: operatorType));
        using (var package = new ExcelPackage(stream))
            ((ExcelDataValidationDecimal)package.Workbook.Worksheets[0].DataValidations[0]).Operator.Should().Be(expected);
    }
}