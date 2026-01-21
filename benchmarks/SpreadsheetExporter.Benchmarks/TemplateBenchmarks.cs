using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Order;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Benchmarks;

[MemoryDiagnoser]
[Orderer(SummaryOrderPolicy.FastestToSlowest)]
[RankColumn]
public class TemplateBenchmarks {
    private List<PersonRecord> recordList = null!;
    private System.Data.DataTable dataTable = null!;

    [Params(100, 1000, 10000)]
    public int RowCount;

    [GlobalSetup]
    public void Setup() {
        recordList = [];
        dataTable = new System.Data.DataTable();
        dataTable.Columns.Add("Id", typeof(int));
        dataTable.Columns.Add("Name", typeof(string));
        dataTable.Columns.Add("Age", typeof(int));
        dataTable.Columns.Add("Email", typeof(string));
        dataTable.Columns.Add("Salary", typeof(decimal));

        for (int i = 0; i < RowCount; i++) {
            recordList.Add(new PersonRecord {
                Id = i,
                Name = $"Person{i}",
                Age = 20 + (i % 50),
                Email = $"person{i}@example.com",
                Salary = 30000 + (i * 100)
            });

            dataTable.Rows.Add(i, $"Person{i}", 20 + (i % 50), $"person{i}@example.com", 30000 + (i * 100));
        }
    }

    [Benchmark]
    public TemplateContext RecordSetTemplate_GetContext() {
        RecordSetTemplate<PersonRecord> template = new(recordList);
        template.Columns.Add("編號", x => x.Id);
        template.Columns.Add("姓名", x => x.Name);
        template.Columns.Add("年齡", x => x.Age);
        template.Columns.Add("電子郵件", x => x.Email);
        template.Columns.Add("薪資", x => x.Salary);

        return template.GetContext();
    }

    [Benchmark]
    public TemplateContext DataTableTemplate_GetContext() {
        DataTableTemplate template = new(dataTable);
        return template.GetContext();
    }

    [Benchmark]
    public TemplateContext GridTemplate_GetContext() {
        GridTemplate template = new();

        template.CreateRow()
            .CreateCell("編號")
            .CreateCell("姓名")
            .CreateCell("年齡")
            .CreateCell("電子郵件")
            .CreateCell("薪資");

        foreach (PersonRecord record in recordList) {
            template.CreateRow()
                .CreateCell(record.Id)
                .CreateCell(record.Name)
                .CreateCell(record.Age)
                .CreateCell(record.Email)
                .CreateCell(record.Salary);
        }

        return template.GetContext();
    }

    private class PersonRecord {
        public int Id { get; set; }

        public string Name { get; set; } = "";

        public int Age { get; set; }

        public string Email { get; set; } = "";

        public decimal Salary { get; set; }
    }
}
