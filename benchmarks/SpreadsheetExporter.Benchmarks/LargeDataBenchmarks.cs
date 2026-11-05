using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Benchmarks;

/// <summary>
/// Measures field value resolution and export costs for large record sets.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[RankColumn]
public class LargeDataBenchmarks {
    private IReadOnlyList<LargeRecord> records = [];

    /// <summary>
    /// Gets or sets the number of records in the data source.
    /// </summary>
    [Params(50_000, 200_000)]
    public int RecordCount { get; set; }

    /// <summary>
    /// Prepares deterministic benchmark data and registers the ClosedXML renderer.
    /// </summary>
    [GlobalSetup]
    public void Setup() {
        SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
        records = Enumerable.Range(1, RecordCount)
            .Select(index => new LargeRecord(
                index,
                $"Customer {index:000000}",
                new LargeAddress($"City {index % 100}", $"{index % 1000:000}"),
                Math.Round(100m + (index * 0.73m), 2),
                DateTime.UnixEpoch.AddDays(index % 3650)
            ))
            .ToList()
            .AsReadOnly();
    }

    /// <summary>
    /// Resolves every cell value and style of a record set bound with property expressions, without rendering.
    /// </summary>
    /// <returns>The number of resolved non-null values.</returns>
    [Benchmark(Baseline = true)]
    public int ResolveExpressionColumns() {
        return ResolveCells(CreateExpressionTemplate());
    }

    /// <summary>
    /// Resolves every cell value and style of a record set bound with string field keys, without rendering.
    /// </summary>
    /// <returns>The number of resolved non-null values.</returns>
    [Benchmark]
    public int ResolveFieldKeyColumns() {
        return ResolveCells(CreateFieldKeyTemplate());
    }

    /// <summary>
    /// Exports a large record set to a byte array.
    /// </summary>
    /// <returns>The generated workbook bytes.</returns>
    [Benchmark]
    public byte[] ExportToByteArray() {
        return CreateDocument().Export();
    }

    /// <summary>
    /// Exports a large record set directly to a stream.
    /// </summary>
    [Benchmark]
    public void ExportToStream() {
        CreateDocument().Export(Stream.Null);
    }

    private SpreadsheetDocument CreateDocument() {
        SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
        document
            .CreateSheet("Orders")
            .SetAutoFilterEnabled()
            .AddTemplate(CreateExpressionTemplate());
        return document;
    }

    private RecordSetTemplate<LargeRecord> CreateExpressionTemplate() {
        RecordSetTemplate<LargeRecord> template = new(records);
        template.Columns
            .Add("Id", x => x.Id, 10)
            .Add("Customer", x => x.Customer, 24)
            .Add("City", x => x.Address.City, 16)
            .Add("Zip", x => x.Address.Zip, 8)
            .Add("Amount", x => x.Amount, 14)
            .Add("OrderDate", x => x.OrderDate, 14);
        return template;
    }

    private RecordSetTemplate<LargeRecord> CreateFieldKeyTemplate() {
        RecordSetTemplate<LargeRecord> template = new(records);
        template.Columns
            .Add<int>("Id", nameof(LargeRecord.Id), 10)
            .Add<string>("Customer", nameof(LargeRecord.Customer), 24)
            .Add<string>("City", "Address.City", 16)
            .Add<string>("Zip", "Address.Zip", 8)
            .Add<decimal>("Amount", nameof(LargeRecord.Amount), 14)
            .Add<DateTime>("OrderDate", nameof(LargeRecord.OrderDate), 14);
        return template;
    }

    private static int ResolveCells(ISheetTemplate template) {
        int count = 0;
        foreach (Cell cell in template.GetLayout().Cells) {
            if (cell.GetValue() is not null) {
                count++;
            }

            cell.GetCellStyle();
        }

        return count;
    }

    private sealed record LargeAddress(string City, string Zip);

    private sealed record LargeRecord(int Id, string Customer, LargeAddress Address, decimal Amount, DateTime OrderDate);
}
