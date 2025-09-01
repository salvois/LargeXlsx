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
using System.Text;
using System.Xml;
using NUnit.Framework;
using Shouldly;

namespace LargeXlsx.Tests;

public static class CustomWriterTest
{
    [Test]
    public static void AppendEscapedXmlText_AllValid()
    {
        using var memoryStream = new MemoryStream();
        new CustomWriter().AppendEscapedXmlText("Lorem 'ipsum' & \"dolor\" \U0001d11e <sit> amet", skipInvalidCharacters: false).FlushTo(memoryStream);
        Encoding.UTF8.GetString(memoryStream.ToArray()).ShouldBe("Lorem 'ipsum' &amp; \"dolor\" \U0001d11e &lt;sit&gt; amet");
    }

    [TestCase(new[] { 'a', '\0', 'b' })]
    [TestCase(new[] { 'a', '\ud800', 'b' })]
    [TestCase(new[] { 'a', '\udc00', 'b' })]
    public static void AppendEscapedXmlText_InvalidChars_Throw(char[] value) =>
        Should.Throw<XmlException>(() => new CustomWriter().AppendEscapedXmlText(new string(value), skipInvalidCharacters: false));

    [TestCase(new[] { 'a', '\0', 'b' })]
    [TestCase(new[] { 'a', '\ud800', 'b' })]
    [TestCase(new[] { 'a', '\udc00', 'b' })]
    public static void AppendEscapedXmlText_InvalidChars_Skip(char[] value)
    {
        using var memoryStream = new MemoryStream();
        new CustomWriter().AppendEscapedXmlText(new string(value), skipInvalidCharacters: true).FlushTo(memoryStream);
        Encoding.UTF8.GetString(memoryStream.ToArray()).ShouldBe("ab");
    }

    [Test]
    public static void AppendEscapedXmlAttribute_AllValid()
    {
        using var memoryStream = new MemoryStream();
        new CustomWriter().AppendEscapedXmlAttribute("Lorem 'ipsum' & \"dolor\" \U0001d11e <sit> amet", skipInvalidCharacters: false).FlushTo(memoryStream);
        Encoding.UTF8.GetString(memoryStream.ToArray()).ShouldBe("Lorem &apos;ipsum&apos; &amp; &quot;dolor&quot; \U0001d11e &lt;sit&gt; amet");
    }

    [TestCase(new[] { 'a', '\0', 'b' })]
    [TestCase(new[] { 'a', '\ud800', 'b' })]
    [TestCase(new[] { 'a', '\udc00', 'b' })]
    public static void AppendEscapedXmlAttribute_InvalidChars_Throw(char[] value) =>
        Should.Throw<XmlException>(() => new CustomWriter().AppendEscapedXmlAttribute(new string(value), skipInvalidCharacters: false));

    [TestCase(new[] { 'a', '\0', 'b' })]
    [TestCase(new[] { 'a', '\ud800', 'b' })]
    [TestCase(new[] { 'a', '\udc00', 'b' })]
    public static void AppendEscapedXmlAttribute_InvalidChars_Skip(char[] value)
    {
        using var memoryStream = new MemoryStream();
        new CustomWriter().AppendEscapedXmlAttribute(new string(value), skipInvalidCharacters: true).FlushTo(memoryStream);
        Encoding.UTF8.GetString(memoryStream.ToArray()).ShouldBe("ab");
    }

    [TestCase(" leading space", " xml:space=\"preserve\"")]
    [TestCase("\tleading tab", " xml:space=\"preserve\"")]
    [TestCase("\rleading CR", " xml:space=\"preserve\"")]
    [TestCase("\nleading LF", " xml:space=\"preserve\"")]
    [TestCase("trailing space ", " xml:space=\"preserve\"")]
    [TestCase("trailing tab\t", " xml:space=\"preserve\"")]
    [TestCase("trailing CR\n", " xml:space=\"preserve\"")]
    [TestCase("trailing LF\n", " xml:space=\"preserve\"")]
    [TestCase("middle   spaces", "")]
    public static void AddSpacePreserveIfNeeded_LeadingOrTrailingWhitespace(string value, string expectation)
    {
        using var memoryStream = new MemoryStream();
        new CustomWriter().AddSpacePreserveIfNeeded(value).FlushTo(memoryStream);
        Encoding.UTF8.GetString(memoryStream.ToArray()).ShouldBe(expectation);
    }
}