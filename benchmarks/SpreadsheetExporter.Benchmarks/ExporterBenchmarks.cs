using System.Drawing;
using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[RankColumn]
public class ExporterBenchmarks {
    private GridTemplate template = null!;
    private readonly string tempDirectory = Path.Combine(Path.GetTempPath(), "SpreadsheetExporter.Benchmarks");

    [Params(100, 1000, 10000)]
    public int RowCount;

    [GlobalSetup]
    public void Setup() {
        Directory.CreateDirectory(tempDirectory);

        template = new GridTemplate();

        template.CreateRow()
            .CreateCell("編號")
            .CreateCell("姓名")
            .CreateCell("年齡")
            .CreateCell("電子郵件")
            .CreateCell("薪資");

        for (int i = 0; i < RowCount; i++) {
            template.CreateRow()
                .CreateCell(i)
                .CreateCell($"Person{i}")
                .CreateCell(20 + (i % 50))
                .CreateCell($"person{i}@example.com")
                .CreateCell(30000 + (i * 100));
        }
    }

    [GlobalCleanup]
    public void Cleanup() {
        if (Directory.Exists(tempDirectory)) {
            Directory.Delete(tempDirectory, true);
        }
    }

    [Benchmark]
    public void NPOI_Export() {
        Excel.NPOI.ExcelExporter exporter = new();
        Sheeter sheeter = exporter.CreateSheeter("測試");
        sheeter.AddTemplate(template);

        byte[] bytes = exporter.Export();
        string filePath = Path.Combine(tempDirectory, $"npoi_{RowCount}.xlsx");
        File.WriteAllBytes(filePath, bytes);
    }

    [Benchmark]
    public void EPPlus_Export() {
        Excel.EPPlus.ExcelExporter exporter = new();
        Sheeter sheeter = exporter.CreateSheeter("測試");
        sheeter.AddTemplate(template);

        byte[] bytes = exporter.Export();
        string filePath = Path.Combine(tempDirectory, $"epplus_{RowCount}.xlsx");
        File.WriteAllBytes(filePath, bytes);
    }

    [Benchmark]
    public void NPOI_Export_WithStyles() {
        Excel.NPOI.ExcelExporter exporter = new();
        Sheeter sheeter = exporter.CreateSheeter("測試");

        GridTemplate styledTemplate = CreateStyledTemplate();
        sheeter.AddTemplate(styledTemplate);

        byte[] bytes = exporter.Export();
        string filePath = Path.Combine(tempDirectory, $"npoi_styled_{RowCount}.xlsx");
        File.WriteAllBytes(filePath, bytes);
    }

    [Benchmark]
    public void EPPlus_Export_WithStyles() {
        Excel.EPPlus.ExcelExporter exporter = new();
        Sheeter sheeter = exporter.CreateSheeter("測試");

        GridTemplate styledTemplate = CreateStyledTemplate();
        sheeter.AddTemplate(styledTemplate);

        byte[] bytes = exporter.Export();
        string filePath = Path.Combine(tempDirectory, $"epplus_styled_{RowCount}.xlsx");
        File.WriteAllBytes(filePath, bytes);
    }

    private GridTemplate CreateStyledTemplate() {
        GridTemplate styledTemplate = new();

        CellStyle headerStyle = new() {
            BackgroundColor = Color.LightBlue,
            Font = new CellFont(Style: FontStyles.IsBold),
            HorizontalAlignment = HorizontalAlignment.Center
        };

        styledTemplate.CreateRow()
            .CreateCell("編號", cellStyle: headerStyle)
            .CreateCell("姓名", cellStyle: headerStyle)
            .CreateCell("年齡", cellStyle: headerStyle)
            .CreateCell("電子郵件", cellStyle: headerStyle)
            .CreateCell("薪資", cellStyle: headerStyle);

        CellStyle dataStyle = new() {
            HorizontalAlignment = HorizontalAlignment.Left
        };

        for (int i = 0; i < RowCount; i++) {
            styledTemplate.CreateRow()
                .CreateCell(i, cellStyle: dataStyle)
                .CreateCell($"Person{i}", cellStyle: dataStyle)
                .CreateCell(20 + (i % 50), cellStyle: dataStyle)
                .CreateCell($"person{i}@example.com", cellStyle: dataStyle)
                .CreateCell(30000 + (i * 100), cellStyle: dataStyle);
        }

        return styledTemplate;
    }
}
