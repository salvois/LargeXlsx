/*
LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2024 Salvatore ISAJA. All rights reserved.

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
using System.Collections.Generic;

namespace LargeXlsx
{
    public class XlsxNumberFormat : IEquatable<XlsxNumberFormat>
    {
        public static readonly XlsxNumberFormat General = new XlsxNumberFormat("general");
        public static readonly XlsxNumberFormat Integer = new XlsxNumberFormat("0");
        public static readonly XlsxNumberFormat TwoDecimal = new XlsxNumberFormat("0.00");
        public static readonly XlsxNumberFormat ThousandInteger = new XlsxNumberFormat("#,##0");
        public static readonly XlsxNumberFormat ThousandTwoDecimal = new XlsxNumberFormat("#,##0.00");
        public static readonly XlsxNumberFormat IntegerPercentage = new XlsxNumberFormat("0%");
        public static readonly XlsxNumberFormat TwoDecimalPercentage = new XlsxNumberFormat("0.00%");
        public static readonly XlsxNumberFormat Scientific = new XlsxNumberFormat("0.00E+00");
        public static readonly XlsxNumberFormat ShortDate = new XlsxNumberFormat("dd/mm/yyyy");
        public static readonly XlsxNumberFormat ShortDateTime = new XlsxNumberFormat("dd/mm/yyyy hh:mm");
        public static readonly XlsxNumberFormat Text = new XlsxNumberFormat("@");
        
        public string FormatCode { get; }

        public XlsxNumberFormat(string formatCode)
        {
            FormatCode = formatCode;
        }

        public override bool Equals(object obj)
        {
            return Equals(obj as XlsxNumberFormat);
        }

        public bool Equals(XlsxNumberFormat other)
        {
            return other != null && FormatCode == other.FormatCode;
        }

        public override int GetHashCode()
        {
            return FormatCode.GetHashCode();
        }

        public static bool operator ==(XlsxNumberFormat format1, XlsxNumberFormat format2)
        {
            return EqualityComparer<XlsxNumberFormat>.Default.Equals(format1, format2);
        }

        public static bool operator !=(XlsxNumberFormat format1, XlsxNumberFormat format2)
        {
            return !(format1 == format2);
        }
    }
}