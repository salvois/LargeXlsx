using System;
using System.IO;
using System.Runtime.InteropServices;
using FluentAssertions;
using NUnit.Framework;
using SharpCompress.Compressors.Deflate;

namespace LargeXlsx.Tests;

[TestFixture]
public static class DisposeTest
{
    [Test]
    public static void DisposeWriterTwice()
    {
        using var stream = new MemoryStream();

        var xlsxWriter = new XlsxWriter(stream);
        xlsxWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Hello World!");
        xlsxWriter.Dispose();

        // dispose a 2nd time (should be idempotent)
        Assert.DoesNotThrow(() => xlsxWriter.Dispose());
    }

    [Test]
    public static void CanDeriveFromWriterSafely()
    {
        using var stream = new MemoryStream();

        var myWriter = new MyWriter(stream);
        myWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Hello World!");
        myWriter.Dispose();

        Assert.True(myWriter.IsMyWriterDisposed);
        Assert.True(myWriter.IsDisposed); // base is disposed too
    }

    [Test]
    public static void FinalizerCleansBase()
    {
        WeakReference? weak = null;

        // track with WeakReference
        var action = new Action(() =>
        {
            using var stream = new MemoryStream();
            var myWriter = new MyWriter(stream);
            myWriter.BeginWorksheet("Sheet 1").BeginRow().Write("Hello World!");
            myWriter.IsMyWriterDisposed.Should().Be(false);
            weak = new WeakReference(myWriter, true); // true = Track reference _after_ Finalize()
        });

        action();

        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        // finalizers call dispose

        ((MyWriter?)weak!.Target)?.IsMyWriterDisposed.Should().Be(true);
        ((MyWriter?)weak!.Target)?.IsDisposed.Should().Be(true);
    }

    class MyWriter : XlsxWriter
    {
        private bool _disposed;

        // sample unmanaged object
        private readonly IntPtr _unmanagedPointer;

        // sample managed object
        private readonly MemoryStream _memoryStream;

        public MyWriter(Stream stream, CompressionLevel compressionLevel = CompressionLevel.Level3, bool useZip64 = false) 
            : base(stream, compressionLevel, useZip64)
        {
            _unmanagedPointer = Marshal.AllocHGlobal(1024);
            _memoryStream = new MemoryStream();
        }

        public bool IsMyWriterDisposed => _disposed;

        protected override void Dispose(bool disposing)
        {
            if (_disposed)
            {
                return;
            }

            if (disposing)
            {
                _memoryStream.Dispose();
            }

            Marshal.FreeHGlobal(_unmanagedPointer);
            _disposed = true;

            base.Dispose(disposing);
        }

        ~MyWriter()
        {
            Dispose(false);
        }
    }
}