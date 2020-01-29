using System;
using System.Globalization;
using System.IO;
using System.Text;

namespace LargeXlsx
{
    internal class InvariantCultureStreamWriter : StreamWriter
    {
        public InvariantCultureStreamWriter(Stream stream) : base(stream, Encoding.UTF8) { }
        public override IFormatProvider FormatProvider => CultureInfo.InvariantCulture;
    }
}