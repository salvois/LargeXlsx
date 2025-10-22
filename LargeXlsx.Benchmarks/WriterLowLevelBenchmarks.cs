using BenchmarkDotNet.Attributes;

namespace LargeXlsx.Benchmarks;

[MemoryDiagnoser]
public class WriterLowLevelBenchmarks
{
    private List<string>? _stringValues;

    [Params(1000, 10000)]
    public int Rows;

    [Params(4, 16)]
    public int Cols;

    private static MemoryStream CreateMemoryStream()
    {
        return new MemoryStream(2 * 1024 * 1024);
    }

    private static XlsxWriter CreateWriter(MemoryStream memoryStream)
    {
        return new XlsxWriter(memoryStream);
    }

    private static void CreateWorksheet(XlsxWriter writer)
    {
        writer.BeginWorksheet("Sheet1");
    }

    [GlobalSetup]
    public void Setup()
    {
        _stringValues = new List<string>(Rows * Cols);
        for (var i = 0; i < Rows * Cols; i++)
            _stringValues.Add($"Value_{i}");
    }

    [GlobalCleanup]
    public void Cleanup()
    {
    }

    [Benchmark]
    public void Write_Int_DefaultStyle()
    {
        using var stream = CreateMemoryStream();
        using var writer = CreateWriter(stream);
        CreateWorksheet(writer);

        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            for (var c = 0; c < Cols; c++)
                writer.Write(r * Cols + c);
        }
    }

    [Benchmark]
    public void Write_Double_DefaultStyle()
    {
        using var stream = CreateMemoryStream();
        using var writer = CreateWriter(stream);
        CreateWorksheet(writer);

        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            for (var c = 0; c < Cols; c++)
                writer.Write(r * Cols + c + 0.5);
        }
    }

    [Benchmark]
    public void Write_String_Inline()
    {
        ArgumentNullException.ThrowIfNull(_stringValues);

        using var stream = CreateMemoryStream();
        using var writer = CreateWriter(stream);
        CreateWorksheet(writer);

        var idx = 0;
        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            for (var c = 0; c < Cols; c++)
                writer.Write(_stringValues[idx++]);
        }
    }

    [Benchmark]
    public void Write_String_Shared()
    {
       ArgumentNullException.ThrowIfNull(_stringValues);

       using var stream = CreateMemoryStream();
       using var writer = CreateWriter(stream);
       CreateWorksheet(writer);

        var idx = 0;
        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            for (var c = 0; c < Cols; c++)
                writer.WriteSharedString(_stringValues[idx++]);
        }
    }

    [Benchmark]
    public void Write_Formula_WithResult()
    {
        using var stream = CreateMemoryStream();
        using var writer = CreateWriter(stream);
        CreateWorksheet(writer);

        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            for (var c = 0; c < Cols; c++)
                writer.WriteFormula($"A{r+1}+B{c+1}", result: r + c);
        }
    }

    [Benchmark]
    public void Write_MergedCells_ColumnSpan()
    {
        using var stream = CreateMemoryStream();
        using var writer = CreateWriter(stream);
        CreateWorksheet(writer);

        for (var r = 0; r < Rows; r++)
        {
            writer.BeginRow();
            writer.Write("Merged", columnSpan: Cols);
        }
    }
}