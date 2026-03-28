using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Benchmarks;

/// <summary>
/// Measures core template preparation costs without focusing on renderer internals.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[RankColumn]
public class TemplateBenchmarks {
    private IReadOnlyList<BenchmarkRecord> records = [];
    private string jsonPayload = "";

    /// <summary>
    /// Gets or sets the number of records used to prepare template contexts.
    /// </summary>
    [Params(100, 1_000, 5_000)]
    public int RecordCount { get; set; }

    /// <summary>
    /// Prepares reusable benchmark data and JSON payloads.
    /// </summary>
    [GlobalSetup]
    public void Setup() {
        SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());

        records = Enumerable.Range(1, RecordCount)
            .Select(index => new BenchmarkRecord(
                index,
                $"Item {index:0000}",
                (index % 25) + 1,
                Math.Round(10m + (index * 0.42m), 2)
            ))
            .ToList()
            .AsReadOnly();

        jsonPayload = CreateJsonPayload(records);
    }

    /// <summary>
    /// Builds a grid template context similar to a lightweight dashboard sheet.
    /// </summary>
    /// <returns>The number of generated cells.</returns>
    [Benchmark]
    public int BuildGridTemplateContext() {
        GridTemplate template = new();

        foreach (BenchmarkRecord record in records) {
            template.CreateRow()
                .CreateCell(record.Code)
                .CreateCell(record.Quantity)
                .CreateCell(record.Amount);
        }

        return template.GetLayout().Cells.Count;
    }

    /// <summary>
    /// Builds a record-set template context using the column model.
    /// </summary>
    /// <returns>The number of generated cells.</returns>
    [Benchmark(Baseline = true)]
    public int BuildRecordSetTemplateContext() {
        RecordSetTemplate<BenchmarkRecord> template = new(records);

        template.Columns
            .Add("ID", static item => item.Id)
            .Add("Code", static item => item.Code)
            .Add("Quantity", static item => item.Quantity)
            .Add("Amount", static item => item.Amount);

        return template.GetLayout().Cells.Count;
    }

    /// <summary>
    /// Parses a JSON workbook definition through <see cref="SpreadsheetDocument.FromJson(string)"/>.
    /// </summary>
    /// <returns>The number of templates on the first parsed sheet.</returns>
    [Benchmark]
    public int ParseJsonDocument() {
        SpreadsheetDocument document = SpreadsheetDocument.FromJson(jsonPayload);
        return document.GetSheet(0).Templates.Count;
    }

    /// <summary>
    /// Parses a JSON workbook definition and resolves the first sheet layout.
    /// </summary>
    /// <returns>The number of generated cells on the first parsed sheet.</returns>
    [Benchmark]
    public int ParseJsonAndBuildLayout() {
        SpreadsheetDocument document = SpreadsheetDocument.FromJson(jsonPayload);
        SheetDefinition sheet = document.GetSheet(0);
        SheetLayout layout = new(sheet);
        return layout.Cells.Count;
    }

    /// <summary>
    /// Parses a JSON workbook definition and exports the workbook.
    /// </summary>
    /// <returns>The generated workbook bytes.</returns>
    [Benchmark]
    public byte[] ParseJsonAndExportWorkbook() {
        SpreadsheetDocument document = SpreadsheetDocument.FromJson(jsonPayload);
        return document.Export();
    }

    private static string CreateJsonPayload(IEnumerable<BenchmarkRecord> source) {
        string recordsJson = string.Join(
            "," + Environment.NewLine,
            source.Select(record =>
                $$"""
                { "Id": {{record.Id}}, "Code": "{{record.Code}}", "Quantity": {{record.Quantity}}, "Amount": {{record.Amount}} }
                """
            ));

        return $$"""
            [
              {
                "SheetName": "Orders",
                "FreezePanes": { "Row": 1, "Column": 0 },
                "IsAutoFilterEnabled": true,
                "Templates": [
                  {
                    "Type": "RecordSet",
                    "Columns": [
                      { "HeaderText": "ID", "FieldKey": "Id" },
                      { "HeaderText": "Code", "FieldKey": "Code" },
                      { "HeaderText": "Quantity", "FieldKey": "Quantity" },
                      { "HeaderText": "Amount", "FieldKey": "Amount" }
                    ],
                    "Records": [
            {{recordsJson}}
                    ]
                  }
                ]
              }
            ]
            """;
    }

    private sealed record BenchmarkRecord(int Id, string Code, int Quantity, decimal Amount);
}
