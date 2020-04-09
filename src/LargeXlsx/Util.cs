using System;
using System.Text;

namespace LargeXlsx
{
    internal static class Util
    {
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
            var columnName = new StringBuilder(3);
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
    }
}