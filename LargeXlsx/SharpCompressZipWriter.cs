using System;
using SharpCompress.Common;
using SharpCompress.Writers;
using SharpCompress.Writers.Zip;
using System.IO;
using SharpCompress.Compressors.Deflate;

namespace LargeXlsx;

public class SharpCompressZipWriter : IDisposable
{
    private readonly ZipWriter _zipWriter;

    public SharpCompressZipWriter(Stream stream, XlsxCompressionLevel compressionLevel, bool useZip64)
    {
        var deflateCompressionLevel = compressionLevel switch
        {
            XlsxCompressionLevel.Fastest => CompressionLevel.BestSpeed,
            XlsxCompressionLevel.Excel => CompressionLevel.Level3,
            XlsxCompressionLevel.Optimal => CompressionLevel.Default,
            XlsxCompressionLevel.Best => CompressionLevel.BestCompression,
            _ => throw new ArgumentOutOfRangeException(nameof(compressionLevel), compressionLevel, null)
        };
        _zipWriter = (ZipWriter)WriterFactory.Open(stream, ArchiveType.Zip, new ZipWriterOptions(CompressionType.Deflate)
        {
            DeflateCompressionLevel = deflateCompressionLevel,
            UseZip64 = useZip64
        });
    }

    public Stream CreateEntry(string path) =>
        _zipWriter.WriteToStream(path, new ZipWriterEntryOptions());

    public void Dispose() =>
        _zipWriter.Dispose();
}