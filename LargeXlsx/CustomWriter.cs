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
using System;
using System.Globalization;
using System.IO;
using System.Text;
using System.Xml;

namespace LargeXlsx;

internal sealed class CustomWriter(int capacityInBytes = 0)
{
    private readonly char[] _charBuffer = new char[1024];
    private readonly byte[] _byteBuffer = new byte[4096];
    private readonly Encoder _encoder = Encoding.UTF8.GetEncoder();
    private readonly MemoryStream _memoryStream = new(capacityInBytes);

    public CustomWriter Append(double value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, charsWritten, _byteBuffer, 0, true);
#else
        var s = value.ToString(CultureInfo.InvariantCulture);
        s.CopyTo(0, _charBuffer, 0, s.Length);
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, s.Length, _byteBuffer, 0, true);
#endif
        _memoryStream.Write(_byteBuffer, 0, bytesWritten);
        return this;
    }

    public CustomWriter Append(decimal value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, charsWritten, _byteBuffer, 0, true);
#else
        var s = value.ToString(CultureInfo.InvariantCulture);
        s.CopyTo(0, _charBuffer, 0, s.Length);
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, s.Length, _byteBuffer, 0, true);
#endif
        _memoryStream.Write(_byteBuffer, 0, bytesWritten);
        return this;
    }

    public CustomWriter Append(int value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, charsWritten, _byteBuffer, 0, true);
#else
        var s = value.ToString(CultureInfo.InvariantCulture);
        s.CopyTo(0, _charBuffer, 0, s.Length);
        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, s.Length, _byteBuffer, 0, true);
#endif
        _memoryStream.Write(_byteBuffer, 0, bytesWritten);
        return this;
    }

    public CustomWriter Append(ReadOnlySpan<byte> value)
    {
#if NETCOREAPP2_1_OR_GREATER
        _memoryStream.Write(value);
#else
        for (var i = 0; i < value.Length; i += _byteBuffer.Length)
        {
            var count = Math.Min(value.Length - i, _byteBuffer.Length);
            value.Slice(i, count).CopyTo(_byteBuffer);
            _memoryStream.Write(_byteBuffer, i, count);
        }
#endif
        return this;
    }

    public CustomWriter Append(byte[] buffer, int offset, int count)
    {
        _memoryStream.Write(buffer, offset, count);
        return this;
    }

    public CustomWriter AppendEscapedXmlText(string value, bool skipInvalidCharacters)
    {
        // A plain old for provides a measurable improvement on garbage collection
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (XmlConvert.IsXmlChar(c))
            {
                if (c == '<') Append("&lt;"u8);
                else if (c == '>') Append("&gt;"u8);
                else if (c == '&') Append("&amp;"u8);
                else
                {
                    _charBuffer[0] = c;
                    var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 1, _byteBuffer, 0, true);
                    _memoryStream.Write(_byteBuffer, 0, bytesWritten);
                }
            }
            else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
            {
                _charBuffer[0] = c;
                _charBuffer[1] = value[i + 1];
                var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 2, _byteBuffer, 0, true);
                _memoryStream.Write(_byteBuffer, 0, bytesWritten);
                i++;
            }
            else if (!skipInvalidCharacters)
                throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
        }
        return this;
    }

    public CustomWriter AppendEscapedXmlAttribute(string value, bool skipInvalidCharacters)
    {
        // A plain old for provides a measurable improvement on garbage collection
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            if (XmlConvert.IsXmlChar(c))
            {
                if (c == '<') Append("&lt;"u8);
                else if (c == '>') Append("&gt;"u8);
                else if (c == '&') Append("&amp;"u8);
                else if (c == '\'') Append("&apos;"u8);
                else if (c == '"') Append("&quot;"u8);
                else
                {
                    _charBuffer[0] = c;
                    var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 1, _byteBuffer, 0, true);
                    _memoryStream.Write(_byteBuffer, 0, bytesWritten);
                }
            }
            else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
            {
                _charBuffer[0] = c;
                _charBuffer[1] = value[i + 1];
                var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 2, _byteBuffer, 0, true);
                _memoryStream.Write(_byteBuffer, 0, bytesWritten);
                i++;
            }
            else if (!skipInvalidCharacters)
                throw new XmlException($"Invalid XML character at position {i} in \"{value}\"");
        }
        return this;
    }

    public CustomWriter AddSpacePreserveIfNeeded(string value)
    {
        if (value.Length > 0 && (XmlConvert.IsWhitespaceChar(value[0]) || XmlConvert.IsWhitespaceChar(value[value.Length - 1])))
            Append(" xml:space=\"preserve\""u8);
        return this;
    }

    public void FlushTo(Stream outputStream)
    {
        _memoryStream.Position = 0;
        _memoryStream.CopyTo(outputStream);
        _memoryStream.SetLength(0);
    }

    public void FlushToIfBiggerThan(Stream outputStream, int byteThreshold)
    {
        if (_memoryStream.Length >= byteThreshold)
            FlushTo(outputStream);
    }

    public int GetUtf8Bytes(int value, byte[] destination)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        return _encoder.GetBytes(_charBuffer, 0, charsWritten, destination, 0, true);
#else
        var s = value.ToString(CultureInfo.InvariantCulture);
        s.CopyTo(0, _charBuffer, 0, s.Length);
        return _encoder.GetBytes(_charBuffer, 0, s.Length, destination, 0, true);
#endif
    }
}