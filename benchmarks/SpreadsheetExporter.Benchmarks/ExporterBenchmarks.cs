using System.Drawing;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Benchmarks;

/// <summary>
/// Measures end-to-end workbook rendering performance for the ClosedXML pipeline.
/// </summary>
[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[RankColumn]
public class ExporterBenchmarks {
    private IReadOnlyList<BenchmarkOrder> orders = [];

    /// <summary>
    /// Gets or sets the number of records rendered into the benchmark workbook.
    /// </summary>
    [Params(100, 1_000, 5_000)]
    public int RecordCount { get; set; }

    /// <summary>
    /// Prepares deterministic benchmark data and registers the ClosedXML renderer.
    /// </summary>
    [GlobalSetup]
    public void Setup() {
        SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());
        orders = Enumerable.Range(1, RecordCount)
            .Select(index => new BenchmarkOrder(
                index,
                $"Customer {index:0000}",
                Regions[index % Regions.Length],
                Math.Round(250m + (index * 1.37m), 2)
            ))
            .ToList()
            .AsReadOnly();
    }

    /// <summary>
    /// Renders a workbook using the defaults with minimal styling overhead.
    /// </summary>
    /// <returns>The generated workbook bytes.</returns>
    [Benchmark(Baseline = true)]
    public byte[] RenderBasicWorkbook() {
        return CreateWorkbook(includeStyles: false).Export();
    }

    /// <summary>
    /// Renders a workbook that exercises styles, data validation, freeze panes, auto filter, and formulas.
    /// </summary>
    /// <returns>The generated workbook bytes.</returns>
    [Benchmark]
    public byte[] RenderStyledWorkbook() {
        return CreateWorkbook(includeStyles: true).Export();
    }

    private SpreadsheetDocument CreateWorkbook(bool includeStyles) {
        SpreadsheetDocument document = SpreadsheetManager.CreateDocument();
        document.DefaultFont = new CellFont("Calibri", 11);

        SheetDefinition overviewSheet = document
            .CreateSheet("Overview", defaultRowHeight: 20)
            .SetColumnWidth(0, 18)
            .SetColumnWidth(1, 22)
            .AddTemplate(CreateOverviewTemplate(includeStyles));

        SheetDefinition ordersSheet = document
            .CreateSheet("Orders", defaultRowHeight: 20)
            .ConfigurePageSettings(x => x.PageOrientation = PageOrientation.Landscape)
            .SetColumnWidth(0, 12)
            .SetColumnWidth(1, 22)
            .SetColumnWidth(2, 14)
            .SetColumnWidth(3, 14)
            .SetColumnWidth(4, 14)
            .SetFreezePanes(0, 1)
            .SetAutoFilterEnabled()
            .AddTemplate(CreateOrdersTemplate(includeStyles));

        return document;
    }

    private GridTemplate CreateOverviewTemplate(bool includeStyles) {
        CellStyle titleStyle = includeStyles
            ? new CellStyle(
                HorizontalAlignment: HorizontalAlignment.Center,
                VerticalAlignment: VerticalAlignment.Middle,
                HasBorder: true,
                BackgroundColor: Color.FromArgb(31, 78, 121),
                Font: new CellFont("Calibri", 13, Color.White, FontStyles.Bold)
            )
            : new CellStyle(HorizontalAlignment: HorizontalAlignment.Center);
        CellStyle labelStyle = includeStyles
            ? new CellStyle(
                HasBorder: true,
                BackgroundColor: Color.FromArgb(221, 235, 247),
                Font: new CellFont("Calibri", 11, Style: FontStyles.Bold)
            )
            : CellStyle.Empty;

        return new GridTemplate()
            .CreateRow(24)
            .CreateCell("SpreadsheetExporter benchmark", columnSpan: 2, cellStyle: titleStyle)
            .CreateRow()
            .CreateCell("Rows", cellStyle: labelStyle)
            .CreateCell(RecordCount, cellStyle: includeStyles ? new CellStyle(HasBorder: true) : CellStyle.Empty)
            .CreateRow()
            .CreateCell("Renderer", cellStyle: labelStyle)
            .CreateCell("ClosedXML", cellStyle: includeStyles ? new CellStyle(HasBorder: true) : CellStyle.Empty);
    }

    private RecordSetTemplate<BenchmarkOrder> CreateOrdersTemplate(bool includeStyles) {
        CellStyle currencyStyle = includeStyles
            ? new CellStyle(HorizontalAlignment: HorizontalAlignment.Right, DataFormat: "#,##0.00")
            : CellStyle.Empty;

        RecordSetTemplate<BenchmarkOrder> template = new(orders) {
            HeaderHeight = 20,
            RecordHeight = 18
        };

        template.Columns
            .Add("Order ID", static order => order.OrderId)
            .Add("Customer", static order => order.Customer)
            .Add(
                "Region",
                static order => order.Region,
                configureGenerators: static config => config.UseDataValidation(
                    _ => new DataValidation {
                        ValidationType = DataValidationType.List,
                        ListItems = Regions
                    }
                )
            )
            .Add(
                "Amount",
                static order => order.Amount,
                fieldStyleGenerator: _ => currencyStyle
            )
            .Add(
                "Margin",
                configureGenerators: static config => config.UseFormula(
                    context => $"D{context.RowIndex + 2}*0.35"
                ),
                fieldStyleGenerator: _ => currencyStyle
            );

        return template;
    }

    private static readonly string[] Regions = ["North", "South", "East", "West"];

    private sealed record BenchmarkOrder(int OrderId, string Customer, string Region, decimal Amount);
}
