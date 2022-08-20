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
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace LargeXlsx
{
    internal static class Util
    {
        private const int MaxColumnCount = 16384;
        private static readonly DateTime ExcelEpoch = new DateTime(1900, 1, 1);
        private static readonly DateTime Date19000301 = new DateTime(1900, 3, 1);
        private static readonly string[] CachedColumnNames = new string[MaxColumnCount];

        public static string EscapeXmlText(string value)
        {
            var sb = new StringBuilder(value.Length);
            foreach (var c in value)
            {
                if (c == '<') sb.Append("&lt;");
                else if (c == '>') sb.Append("&gt;");
                else if (c == '&') sb.Append("&amp;");
                else sb.Append(c);
            }

            return sb.ToString();
        }

        public static string EscapeXmlAttribute(string value)
        {
            var sb = new StringBuilder(value.Length);
            foreach (var c in value)
            {
                if (c == '<') sb.Append("&lt;");
                else if (c == '>') sb.Append("&gt;");
                else if (c == '&') sb.Append("&amp;");
                else if (c == '\'') sb.Append("&apos;");
                else if (c == '"') sb.Append("&quot;");
                else sb.Append(c);
            }

            return sb.ToString();
        }

        public static string GetColumnName(int columnIndex)
        {
            if (columnIndex < 1 || columnIndex > MaxColumnCount)
                throw new InvalidOperationException($"A worksheet can contain at most {MaxColumnCount} columns ({columnIndex} attempted)");
            var columnName = CachedColumnNames[columnIndex - 1];
            if (columnName == null)
            {
                columnName = GetColumnNameInternal(columnIndex);
                CachedColumnNames[columnIndex - 1] = columnName;
            }
            return columnName;
        }

        private static string GetColumnNameInternal(int columnIndex)
        {
            var columnName = new StringBuilder(3); // This has been measured to be faster than string concatenation
            while (true)
            {
                if (columnIndex > 26)
                {
                    columnIndex = Math.DivRem(columnIndex - 1, 26, out var rem);
                    columnName.Insert(0, (char)('A' + rem));
                }
                else
                {
                    columnName.Insert(0, (char)('A' + columnIndex - 1));
                    return columnName.ToString();
                }
            }
        }

        public static double DateToDouble(DateTime date)
        {
            var days = date.Subtract(ExcelEpoch).TotalDays + 1;
            // Excel wrongly assumes that 1900 is a leap year:
            // https://docs.microsoft.com/en-us/office/troubleshoot/excel/wrongly-assumes-1900-is-leap-year
            if (date >= Date19000301) days++;
            return days;
        }

        public static int BoolToInt(bool value) => value ? 1 : 0;

        public static string EnumToAttributeValue<T>(T enumValue)
        {
            var s = enumValue.ToString();
            return char.ToLowerInvariant(s[0]) + s.Substring(1);
        }

        // Hashing procedure courtesy of https://docs.microsoft.com/en-us/openspecs/office_file_formats/ms-offcrypto/1357ea58-646e-4483-92ef-95d718079d6f
        public static byte[] ComputePasswordHash(string password, byte[] saltValue, int spinCount)
        {
            var hasher = new SHA512Managed();
            var hash = hasher.ComputeHash(saltValue.Concat(Encoding.Unicode.GetBytes(password)).ToArray());
            for (var i = 0; i < spinCount; i++)
            {
                var iterator = BitConverter.GetBytes(i);
                if (!BitConverter.IsLittleEndian)
                    Array.Reverse(iterator);
                hash = hasher.ComputeHash(hash.Concat(iterator).ToArray());
            }
            return hash;
        }
    }
}