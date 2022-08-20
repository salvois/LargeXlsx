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
using System;
using FluentAssertions;
using NUnit.Framework;

namespace LargeXlsx.Tests;

[TestFixture]
public static class UtilTest
{
    [Test]
    public static void EscapeXmlText()
    {
        var escapedXmlText = Util.EscapeXmlText("Lorem 'ipsum' & \"dolor\" <sit> amet");
        escapedXmlText.Should().Be("Lorem 'ipsum' &amp; \"dolor\" &lt;sit&gt; amet");
    }

    [Test]
    public static void EscapeXmlAttribute()
    {
        var escapedXmlText = Util.EscapeXmlAttribute("Lorem 'ipsum' & \"dolor\" <sit> amet");
        escapedXmlText.Should().Be("Lorem &apos;ipsum&apos; &amp; &quot;dolor&quot; &lt;sit&gt; amet");
    }

    [TestCase(1, "A")]
    [TestCase(2, "B")]
    [TestCase(26, "Z")]
    [TestCase(27, "AA")]
    [TestCase(28, "AB")]
    [TestCase(52, "AZ")]
    [TestCase(53, "BA")]
    [TestCase(702, "ZZ")]
    [TestCase(703, "AAA")]
    [TestCase(704, "AAB")]
    [TestCase(729, "ABA")]
    [TestCase(1378, "AZZ")]
    [TestCase(1379, "BAA")]
    [TestCase(16384, "XFD")]
    public static void GetColumnName(int index, string expectedName)
    {
        var name = Util.GetColumnName(index);
        name.Should().Be(expectedName);
    }

    [TestCase(-1)]
    [TestCase(0)]
    [TestCase(16385)]
    public static void GetColumnNameOutOfRange(int index)
    {
        Func<string> act = () => Util.GetColumnName(index);
        act.Should().Throw<InvalidOperationException>();
    }

    [TestCase("2020-05-06T18:27:00", 43957.76875)]
    [TestCase("1900-01-01", 1)]
    [TestCase("1900-02-28", 59)]
    [TestCase("1900-03-01", 61)]
    public static void DateToDouble(string dateString, double expected)
    {
        var date = DateTime.Parse(dateString);
        var serialDate = Util.DateToDouble(date);
        serialDate.Should().BeApproximately(expected, 0.000001);
    }

    [Test]
    public static void ComputePasswordHash() =>
        Convert.ToBase64String(Util.ComputePasswordHash(
                password: "Lorem ipsum",
                saltValue: Convert.FromBase64String("5kelhTC7DUqQ5qi78ihM8A=="),
                spinCount: 100000))
            .Should().Be("/dQmPXViT1u/fiTHmjLlP2HqOjYeRKI8W367Qn/Eikv63K8nnZMiyk2Wl9ShdHaBL7y1AeeJq5gxm4bW0ArxYg==");
}