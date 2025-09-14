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

internal sealed class CustomWriter
{
    private readonly char[] _charBuffer = new char[1024];
    private readonly Encoder _encoder = Encoding.UTF8.GetEncoder();
    private byte[] _writeBuffer = new byte[4096];
    private int _writeBufferLength = 0;

    public CustomWriter Append(double value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        WriteAscii(_charBuffer, charsWritten);
#else
        WriteAscii(value.ToString(CultureInfo.InvariantCulture));
#endif
        return this;
    }

    public CustomWriter Append(decimal value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        WriteAscii(_charBuffer, charsWritten);
#else
        WriteAscii(value.ToString(CultureInfo.InvariantCulture));
#endif
        return this;
    }

    public CustomWriter Append(int value)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        WriteAscii(_charBuffer, charsWritten);
#else
        WriteAscii(value.ToString(CultureInfo.InvariantCulture));
#endif
        return this;
    }

    public CustomWriter Append(ReadOnlySpan<byte> value)
    {
        EnsureDeltaCapacity(value.Length);
        value.CopyTo(new Span<byte>(_writeBuffer, _writeBufferLength, value.Length));
        _writeBufferLength += value.Length;
        return this;
    }

    public CustomWriter AppendEscapedXmlText(string value, bool skipInvalidCharacters)
    {
        // A plain old for provides a measurable improvement on garbage collection
        for (var i = 0; i < value.Length; i++)
        {
            var c = value[i];
            EnsureDeltaCapacity(4);
            if (XmlConvert.IsXmlChar(c))
            {
                if (c == '<') Append("&lt;"u8);
                else if (c == '>') Append("&gt;"u8);
                else if (c == '&') Append("&amp;"u8);
                else
                {
                    if (c < 0x80)
                        _writeBuffer[_writeBufferLength++] = (byte)c;
                    else
                    {
                        _charBuffer[0] = c;
                        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 1, _writeBuffer, _writeBufferLength, true);
                        _writeBufferLength += bytesWritten;
                    }
                }
            }
            else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
            {
                _charBuffer[0] = c;
                i++;
                _charBuffer[1] = value[i];
                var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 2, _writeBuffer, _writeBufferLength, true);
                _writeBufferLength += bytesWritten;
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
            EnsureDeltaCapacity(4);
            if (XmlConvert.IsXmlChar(c))
            {
                if (c == '<') Append("&lt;"u8);
                else if (c == '>') Append("&gt;"u8);
                else if (c == '&') Append("&amp;"u8);
                else if (c == '\'') Append("&apos;"u8);
                else if (c == '"') Append("&quot;"u8);
                else
                {
                    if (c < 0x80)
                        _writeBuffer[_writeBufferLength++] = (byte)c;
                    else
                    {
                        _charBuffer[0] = c;
                        var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 1, _writeBuffer, _writeBufferLength, true);
                        _writeBufferLength += bytesWritten;
                    }
                }
            }
            else if (i < value.Length - 1 && XmlConvert.IsXmlSurrogatePair(value[i + 1], c))
            {
                _charBuffer[0] = c;
                i++;
                _charBuffer[1] = value[i];
                var bytesWritten = _encoder.GetBytes(_charBuffer, 0, 2, _writeBuffer, _writeBufferLength, true);
                _writeBufferLength += bytesWritten;
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
        outputStream.Write(_writeBuffer, 0, _writeBufferLength);
        _writeBufferLength = 0;
    }

    public void FlushToIfBiggerThan(Stream outputStream, int byteThreshold)
    {
        if (_writeBufferLength >= byteThreshold)
            FlushTo(outputStream);
    }

    public int GetUtf8Bytes(int value, byte[] destination)
    {
#if NETCOREAPP2_1_OR_GREATER
        if (!value.TryFormat(_charBuffer, out var charsWritten, provider: CultureInfo.InvariantCulture))
            throw new ArgumentException();
        for (var i = 0; i < charsWritten; i++)
            destination[i] = (byte)_charBuffer[i];
        return charsWritten;
#else
        var s = value.ToString(CultureInfo.InvariantCulture);
        for (var i = 0; i < s.Length; i++)
            destination[i] = (byte)s[i];
        return s.Length;
#endif
    }

    private void WriteAscii(char[] chars, int count)
    {
        EnsureDeltaCapacity(count);
        for (var i = 0; i < count; i++)
            _writeBuffer[_writeBufferLength++] = (byte)chars[i];
    }

    private void WriteAscii(string s)
    {
        // Avoid memory allocations by foreach
        EnsureDeltaCapacity(s.Length);
        for (var i = 0; i < s.Length; i++)
            _writeBuffer[_writeBufferLength++] = (byte)s[i];
    }

    private void EnsureDeltaCapacity(int deltaCapacity)
    {
        var newCapacity = _writeBufferLength + deltaCapacity;
        var capacity = _writeBuffer.Length;
        if (capacity < newCapacity)
        {
            do
            {
                if (capacity >= 1E9)
                    throw new InvalidOperationException("Attempting to buffer too much data");
                capacity *= 2;
            } while (capacity < newCapacity);
            Array.Resize(ref _writeBuffer, capacity);
        }
    }
}