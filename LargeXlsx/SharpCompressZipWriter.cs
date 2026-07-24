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
#if NETSTANDARD2_0
using System;
using System.IO;
using SharpCompress.Common;
using SharpCompress.Writers;
using SharpCompress.Writers.Zip;
using SharpCompress.Compressors.Deflate;

namespace LargeXlsx;

public class SharpCompressZipWriter : IZipWriter
{
    private readonly ZipWriter _zipWriter;

    public SharpCompressZipWriter(Stream stream, XlsxCompressionLevel compressionLevel, bool useZip64)
    {
        var deflateCompressionLevel = compressionLevel switch
        {
            XlsxCompressionLevel.Fastest => CompressionLevel.BestSpeed,
            XlsxCompressionLevel.Optimal => CompressionLevel.Default,
            _ => throw new ArgumentOutOfRangeException(nameof(compressionLevel), compressionLevel, null)
        };
        _zipWriter = (ZipWriter)WriterFactory.OpenWriter(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate)
        {
            CompressionLevel = (int)deflateCompressionLevel,
            UseZip64 = useZip64
        });
    }

    public Stream CreateEntry(string path) =>
        _zipWriter.WriteToStream(path, new ZipWriterEntryOptions());

    public void Dispose() =>
        _zipWriter.Dispose();
}
#endif