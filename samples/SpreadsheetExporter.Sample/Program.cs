using System.Data;
using System.Drawing;
using CloudyWing.SpreadsheetExporter.Config;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.DataTable;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.RecordSet;

namespace CloudyWing.SpreadsheetExporter.Samples;

internal class Program {
    static void Main(string[] args) {
        Console.WriteLine("SpreadsheetExporter 功能展示");
        Console.WriteLine("============================\n");

        string baseOutputDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Outputs");
        Directory.CreateDirectory(baseOutputDir);

        Console.WriteLine("=== 使用 NPOI 產生範例 ===\n");
        string npoiOutputDir = Path.Combine(baseOutputDir, "NPOI");
        Directory.CreateDirectory(npoiOutputDir);

        Console.WriteLine("1. 設定 NPOI SpreadsheetManager");
        SetupNpoiSpreadsheetManager();
        GenerateAllExamples(npoiOutputDir);

        Console.WriteLine("\n=== 使用 EPPlus 產生範例 ===\n");
        string epplusOutputDir = Path.Combine(baseOutputDir, "EPPlus");
        Directory.CreateDirectory(epplusOutputDir);

        Console.WriteLine("1. 設定 EPPlus SpreadsheetManager");
        SetupEpplusSpreadsheetManager();
        GenerateAllExamples(epplusOutputDir);

        Console.WriteLine($"\n所有範例檔案已建立在:");
        Console.WriteLine($"  - NPOI: {npoiOutputDir}");
        Console.WriteLine($"  - EPPlus: {epplusOutputDir}");
        Console.WriteLine("\n按任意鍵結束...");
        Console.ReadKey();
    }

    static void GenerateAllExamples(string outputDir) {
        Console.WriteLine("\n2. 建立基本 GridTemplate 範例");
        CreateBasicGridTemplateExample(outputDir);

        Console.WriteLine("\n3. 建立進階 GridTemplate 範例");
        CreateAdvancedGridTemplateExample(outputDir);

        Console.WriteLine("\n4. 建立 RecordSetTemplate 基本範例");
        CreateRecordSetTemplateExample(outputDir);

        Console.WriteLine("\n5. 建立 RecordSetTemplate 進階範例");
        CreateAdvancedRecordSetTemplateExample(outputDir);

        Console.WriteLine("\n6. 建立 RecordSetTemplate 巢狀物件範例");
        CreateNestedRecordSetTemplateExample(outputDir);

        Console.WriteLine("\n7. 建立 DataTableTemplate 範例");
        CreateDataTableTemplateExample(outputDir);

        Console.WriteLine("\n8. 建立多個 Sheet 範例");
        CreateMultiSheetExample(outputDir);

        Console.WriteLine("\n9. 建立自訂 Template 範例");
        CreateCustomTemplateExample(outputDir);

        Console.WriteLine("\n10. 建立 Sheeter 進階功能範例");
        CreateSheeterAdvancedExample(outputDir);

        Console.WriteLine("\n11. 建立 Freeze Panes 和 AutoFilter 範例");
        CreateFreezePanesAndAutoFilterExample(outputDir);

        Console.WriteLine("\n12. 建立 Data Validation 範例");
        CreateDataValidationExample(outputDir);

        Console.WriteLine("\n13. 建立頁面設定範例");
        CreatePageSettingsExample(outputDir);
    }

    static void SetupNpoiSpreadsheetManager() {
        SetupSpreadsheetManager(() => new Excel.NPOI.ExcelExporter(), "NPOI");
    }

    static void SetupEpplusSpreadsheetManager() {
        SetupSpreadsheetManager(() => new Excel.EPPlus.ExcelExporter(), "EPPlus");
    }

    static void SetupSpreadsheetManager(Func<ExporterBase> exporterFactory, string exporterName) {
        SpreadsheetManager.SetExporter(() => {
            ExporterBase exporter = exporterFactory();
            exporter.DefaultBasicSheetName = "工作表";

            exporter.SpreadsheetExportingEvent += (sender, e) => {
                Console.WriteLine($"   - 正在產生 Spreadsheet...");
            };

            exporter.SpreadsheetExportedEvent += (sender, e) => {
                Console.WriteLine($"   - Spreadsheet 產生完成！");
            };

            return exporter;
        });

        SpreadsheetManager.DefaultCellStyles = new((setuper) => {
            CellStyleConfiguration oldCellStyles = SpreadsheetManager.DefaultCellStyles;
            setuper.CellStyle = oldCellStyles.CellStyle;
            setuper.GridCellStyle = oldCellStyles.GridCellStyle;
            setuper.HeaderStyle = oldCellStyles.HeaderStyle with {
                BackgroundColor = Color.LightBlue,
                Font = oldCellStyles.HeaderStyle.Font with {
                    Style = FontStyles.IsBold
                }
            };
        });

        Console.WriteLine($"   - 已設定 {exporterName} Exporter 和預設樣式");
    }

    static void CreateBasicGridTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("基本範例");

        GridTemplate template = new();
        template.CreateRow()
            .CreateCell("姓名", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("年齡", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("城市", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);

        template.CreateRow()
            .CreateCell("張三")
            .CreateCell(30)
            .CreateCell("台北");

        template.CreateRow()
            .CreateCell("李四")
            .CreateCell(25)
            .CreateCell("高雄");

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"01_基本GridTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateAdvancedGridTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("進階範例");

        GridTemplate template = new();
        CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;

        CellStyle titleStyle = cellStyles.CellStyle with {
            HorizontalAlignment = HorizontalAlignment.Center,
            Font = cellStyles.CellStyle.Font with {
                Size = 16,
                Style = FontStyles.IsBold
            },
            BackgroundColor = Color.LightGreen
        };

        template.CreateRow()
            .CreateCell("銷售報表", columnSpan: 4, cellStyle: titleStyle);

        template.CreateRow()
            .CreateCell("產品", cellStyle: cellStyles.HeaderStyle)
            .CreateCell("單價", cellStyle: cellStyles.HeaderStyle)
            .CreateCell("數量", cellStyle: cellStyles.HeaderStyle)
            .CreateCell("小計", cellStyle: cellStyles.HeaderStyle);

        CellStyle dataStyle = cellStyles.CellStyle with {
            HorizontalAlignment = HorizontalAlignment.Right
        };

        template.CreateRow()
            .CreateCell("產品A", cellStyle: cellStyles.CellStyle)
            .CreateCell(100, cellStyle: dataStyle)
            .CreateCell(5, cellStyle: dataStyle)
            .CreateCell((cell, row) => $"B{row + 1}*C{row + 1}", cellStyle: dataStyle);

        template.CreateRow()
            .CreateCell("產品B", cellStyle: cellStyles.CellStyle)
            .CreateCell(200, cellStyle: dataStyle)
            .CreateCell(3, cellStyle: dataStyle)
            .CreateCell((cell, row) => $"B{row + 1}*C{row + 1}", cellStyle: dataStyle);

        CellStyle totalStyle = cellStyles.HeaderStyle with {
            HorizontalAlignment = HorizontalAlignment.Right
        };

        template.CreateRow()
            .CreateCell("總計", columnSpan: 3, cellStyle: cellStyles.HeaderStyle)
            .CreateCell((cell, row) => $"SUM(D3:D4)", cellStyle: totalStyle);

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"02_進階GridTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateRecordSetTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("員工清單");

        List<Employee> employees = [
            new() { Id = 1, Name = "張三", Department = "業務部", Salary = 45000 },
            new() { Id = 2, Name = "李四", Department = "研發部", Salary = 55000 },
            new() { Id = 3, Name = "王五", Department = "業務部", Salary = 48000 },
            new() { Id = 4, Name = "趙六", Department = "人資部", Salary = 42000 }
        ];

        RecordSetTemplate<Employee> template = new(employees);
        template.Columns.Add("員工編號", x => x.Id);
        template.Columns.Add("姓名", x => x.Name);
        template.Columns.Add("部門", x => x.Department);
        template.Columns.Add("薪資", x => x.Salary);

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"03_RecordSetTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateAdvancedRecordSetTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("員工詳細資料");

        List<Employee> employees = [
            new() { Id = 1, Name = "張三", Department = "業務部", Salary = 45000 },
            new() { Id = 2, Name = "李四", Department = "研發部", Salary = 55000 },
            new() { Id = 3, Name = "王五", Department = "業務部", Salary = 48000 },
            new() { Id = 4, Name = "趙六", Department = "人資部", Salary = 42000 }
        ];

        RecordSetTemplate<Employee> template = new(employees);

        template.Columns.Add("員工編號", x => x.Id);
        template.Columns.Add("姓名", x => x.Name);
        template.Columns.Add("部門", x => x.Department);
        template.Columns.Add("薪資", x => x.Salary, x => x.UseValue(y => $"NT$ {y.Value:N0}"));

        template.Columns.Add(
            "薪資等級",
            x => x.Salary,
            x => x.UseValue(y => y.Value >= 50000 ? "高" : "一般")
        );

        template.Columns.Add(
            "完整資訊",
            x => x.UseValue(y => $"{y.Record.Name} ({y.Record.Department})")
        );

        template.Columns.Add(
            "備註",
            x => x.Salary,
            x => x.UseValue(y => y.Value >= 50000 ? "優秀" : "")
        );

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"04_進階RecordSetTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateNestedRecordSetTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("部門資訊");

        List<DepartmentInfo> departments = [
            new() {
                DepartmentName = "研發部",
                Location = "台北",
                Manager = new() { Id = 101, Name = "陳經理", Department = "研發部", Salary = 80000 },
                Employees = [
                    new() { Id = 1, Name = "李小華", Department = "研發部", Salary = 55000 },
                    new() { Id = 2, Name = "陳大明", Department = "研發部", Salary = 60000 },
                    new() { Id = 3, Name = "林小美", Department = "研發部", Salary = 58000 }
                ]
            },
            new() {
                DepartmentName = "業務部",
                Location = "台中",
                Manager = new() { Id = 102, Name = "王經理", Department = "業務部", Salary = 75000 },
                Employees = [
                    new() { Id = 4, Name = "張小三", Department = "業務部", Salary = 48000 },
                    new() { Id = 5, Name = "劉小四", Department = "業務部", Salary = 52000 }
                ]
            },
            new() {
                DepartmentName = "人資部",
                Location = "高雄",
                Manager = new() { Id = 103, Name = "黃經理", Department = "人資部", Salary = 70000 },
                Employees = [
                    new() { Id = 6, Name = "吳小五", Department = "人資部", Salary = 45000 }
                ]
            }
        ];

        RecordSetTemplate<DepartmentInfo> template = new(departments);

        template.Columns.Add("部門基本資訊")
            .AddChildToLast("部門名稱", x => x.DepartmentName)
            .AddChildToLast("所在地", x => x.Location)
            .Add("主管資訊")
            .AddChildToLast("姓名", x => x.UseValue(y => y.Record.Manager.Name))
            .AddChildToLast("薪資", x => x.UseValue(y => $"NT$ {y.Record.Manager.Salary:N0}"))
            .Add("員工統計");

        template.Columns.Last().ChildColumns.Add("人數", x => x.UseValue(y => y.Record.Employees.Count))
            .AddChildToLast("平均薪資", x => x.UseValue(y => $"NT$ {y.Record.Employees.Average(e => e.Salary):N0}"))
            .AddChildToLast("薪資總和", x => x.UseValue(y => $"NT$ {y.Record.Employees.Sum(e => e.Salary):N0}"));

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"05_巢狀RecordSetTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateDataTableTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("產品資料");

        DataTable dataTable = new();
        dataTable.Columns.Add("產品編號", typeof(string));
        dataTable.Columns.Add("產品名稱", typeof(string));
        dataTable.Columns.Add("單價", typeof(decimal));
        dataTable.Columns.Add("庫存", typeof(int));

        dataTable.Rows.Add("P001", "滑鼠", 299, 100);
        dataTable.Rows.Add("P002", "鍵盤", 899, 50);
        dataTable.Rows.Add("P003", "螢幕", 5999, 30);
        dataTable.Rows.Add("P004", "喇叭", 1299, 75);

        DataTableTemplate template = new(dataTable);
        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"06_DataTableTemplate{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateMultiSheetExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();

        Sheeter sheet1 = exporter.CreateSheeter("銷售資料");
        GridTemplate salesTemplate = new();
        salesTemplate.CreateRow()
            .CreateCell("日期", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("金額", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);
        salesTemplate.CreateRow().CreateCell("2024-01-01").CreateCell(10000);
        salesTemplate.CreateRow().CreateCell("2024-01-02").CreateCell(15000);
        sheet1.AddTemplate(salesTemplate);

        Sheeter sheet2 = exporter.CreateSheeter("產品資料");
        GridTemplate productsTemplate = new();
        productsTemplate.CreateRow()
            .CreateCell("產品", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("單價", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);
        productsTemplate.CreateRow().CreateCell("產品A").CreateCell(100);
        productsTemplate.CreateRow().CreateCell("產品B").CreateCell(200);
        sheet2.AddTemplate(productsTemplate);

        Sheeter sheet3 = exporter.CreateSheeter("摘要");
        GridTemplate summaryTemplate = new();
        summaryTemplate.CreateRow().CreateCell("總銷售額", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);
        summaryTemplate.CreateRow().CreateCell(25000);
        sheet3.AddTemplate(summaryTemplate);

        string filePath = Path.Combine(outputDir, $"07_多個Sheet{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)} (包含 3 個 Sheet)");
    }

    static void CreateCustomTemplateExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("自訂範本");

        ReportInfoTemplate infoTemplate = new("月度銷售報表", DateTime.Now);
        sheeter.AddTemplate(infoTemplate);

        GridTemplate dataTemplate = new();
        dataTemplate.CreateRow()
            .CreateCell("項目", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("金額", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);
        dataTemplate.CreateRow().CreateCell("產品銷售").CreateCell(100000);
        dataTemplate.CreateRow().CreateCell("服務收入").CreateCell(50000);
        sheeter.AddTemplate(dataTemplate);

        string filePath = Path.Combine(outputDir, $"08_自訂Template{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateSheeterAdvancedExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("進階功能");

        sheeter.SetColumnWidth(0, 15);
        sheeter.SetColumnWidth(1, 20);
        sheeter.SetColumnWidth(2, 25);

        GridTemplate template = new();

        template.CreateRow(height: 30)
            .CreateCell("欄位A", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("欄位B", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle)
            .CreateCell("欄位C", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);

        template.CreateRow(height: 20)
            .CreateCell("資料1")
            .CreateCell("資料2")
            .CreateCell("資料3");

        template.CreateRow(height: 20)
            .CreateCell("資料4")
            .CreateCell("資料5")
            .CreateCell("資料6");

        sheeter.AddTemplate(template);

        sheeter.Password = "1234";

        string filePath = Path.Combine(outputDir, $"09_Sheeter進階功能{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)} (密碼: 1234)");
    }

    static void CreateFreezePanesAndAutoFilterExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("凍結與篩選");

        List<Product> products = [];
        for (int i = 1; i <= 20; i++) {
            products.Add(new() {
                ProductId = $"P{i:000}",
                ProductName = $"產品{i}",
                Price = 1000 + (i * 100),
                Stock = 50 + (i * 5)
            });
        }

        RecordSetTemplate<Product> template = new(products);
        template.Columns.Add("產品編號", x => x.ProductId);
        template.Columns.Add("產品名稱", x => x.ProductName);
        template.Columns.Add("單價", x => x.Price);
        template.Columns.Add("庫存", x => x.Stock);

        template.IsFreezeHeader = true;
        template.IsAutoFilterEnabled = true;

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"10_凍結與篩選{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreateDataValidationExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("資料驗證");

        GridTemplate template = new();
        CellStyleConfiguration cellStyles = SpreadsheetManager.DefaultCellStyles;

        template.CreateRow()
            .CreateCell("驗證類型", cellStyle: cellStyles.HeaderStyle)
            .CreateCell("請輸入資料", cellStyle: cellStyles.HeaderStyle)
            .CreateCell("說明", cellStyle: cellStyles.HeaderStyle);

        template.CreateRow()
            .CreateCell("整數 (1-100)")
            .CreateCell(cell => {
                cell.DataValidationGenerator = (col, row) => new() {
                    ValidationType = DataValidationType.Integer,
                    Operator = DataValidationOperator.Between,
                    Value1 = 1,
                    Value2 = 100,
                    ErrorTitle = "輸入錯誤",
                    ErrorMessage = "請輸入 1 到 100 之間的整數",
                    IsErrorAlertShown = true
                };
            })
            .CreateCell("請輸入 1 到 100 之間的整數");

        template.CreateRow()
            .CreateCell("小數 (0-1)")
            .CreateCell(cell => {
                cell.DataValidationGenerator = (col, row) => new() {
                    ValidationType = DataValidationType.Decimal,
                    Operator = DataValidationOperator.Between,
                    Value1 = 0.0,
                    Value2 = 1.0,
                    ErrorTitle = "輸入錯誤",
                    ErrorMessage = "請輸入 0 到 1 之間的小數",
                    IsErrorAlertShown = true
                };
            })
            .CreateCell("請輸入 0 到 1 之間的小數");

        template.CreateRow()
            .CreateCell("清單選擇")
            .CreateCell(cell => {
                cell.DataValidationGenerator = (col, row) => new() {
                    ValidationType = DataValidationType.List,
                    ListItems = ["選項A", "選項B", "選項C"],
                    ErrorTitle = "輸入錯誤",
                    ErrorMessage = "請從清單中選擇",
                    IsErrorAlertShown = true
                };
            })
            .CreateCell("請從清單選擇: 選項A, 選項B, 選項C");

        template.CreateRow()
            .CreateCell("日期")
            .CreateCell(cell => {
                cell.DataValidationGenerator = (col, row) => new() {
                    ValidationType = DataValidationType.Date,
                    Operator = DataValidationOperator.GreaterThan,
                    Formula = "TODAY()",
                    ErrorTitle = "輸入錯誤",
                    ErrorMessage = "請輸入今天之後的日期",
                    IsErrorAlertShown = true
                };
            })
            .CreateCell("請輸入今天之後的日期");

        template.CreateRow()
            .CreateCell("文字長度")
            .CreateCell(cell => {
                cell.DataValidationGenerator = (col, row) => new() {
                    ValidationType = DataValidationType.TextLength,
                    Operator = DataValidationOperator.Between,
                    Value1 = 5,
                    Value2 = 10,
                    ErrorTitle = "輸入錯誤",
                    ErrorMessage = "文字長度必須在 5 到 10 之間",
                    IsErrorAlertShown = true
                };
            })
            .CreateCell("請輸入長度 5-10 的文字");

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"11_資料驗證{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)}");
    }

    static void CreatePageSettingsExample(string outputDir) {
        ISpreadsheetExporter exporter = SpreadsheetManager.CreateExporter();
        Sheeter sheeter = exporter.CreateSheeter("頁面設定");

        sheeter.PageSettings.PaperSize = PaperSize.A4;
        sheeter.PageSettings.PageOrientation = PageOrientation.Landscape;

        GridTemplate template = new();

        template.CreateRow();
        for (int i = 0; i < 10; i++) {
            template.CreateCell($"欄位{i + 1}", cellStyle: SpreadsheetManager.DefaultCellStyles.HeaderStyle);
        }

        for (int row = 0; row < 30; row++) {
            template.CreateRow();
            for (int col = 0; col < 10; col++) {
                template.CreateCell($"R{row + 1}C{col + 1}");
            }
        }

        sheeter.AddTemplate(template);

        string filePath = Path.Combine(outputDir, $"12_頁面設定{exporter.FileNameExtension}");
        exporter.ExportFile(filePath);
        Console.WriteLine($"   - 已建立: {Path.GetFileName(filePath)} (A4橫向)");
    }

}

// 範例資料類別
class Employee {
    public int Id { get; set; }

    public string Name { get; set; }

    public string Department { get; set; }

    public decimal Salary { get; set; }
}

class Product {
    public string ProductId { get; set; }

    public string ProductName { get; set; }

    public decimal Price { get; set; }

    public int Stock { get; set; }
}

class DepartmentInfo {
    public string DepartmentName { get; set; }

    public string Location { get; set; }

    public Employee Manager { get; set; }

    public List<Employee> Employees { get; set; }
}

// 自訂 Template 範例
class ReportInfoTemplate : ITemplate {
    private readonly string title;
    private readonly DateTime createdDate;

    public ReportInfoTemplate(string title, DateTime createdDate) {
        this.title = title;
        this.createdDate = createdDate;
    }

    public TemplateContext GetContext() {
        List<Cell> cells = new();
        Dictionary<int, double?> rowHeights = new();

        CellStyle titleStyle = SpreadsheetManager.DefaultCellStyles.CellStyle with {
            Font = SpreadsheetManager.DefaultCellStyles.CellStyle.Font with {
                Size = 18,
                Style = FontStyles.IsBold
            },
            HorizontalAlignment = HorizontalAlignment.Center
        };

        // 標題儲存格
        cells.Add(new() {
            Point = new(0, 0),
            Size = new(3, 1),
            ValueGenerator = (col, row) => title,
            CellStyleGenerator = (col, row) => titleStyle
        });

        rowHeights[0] = 35;

        // 日期儲存格
        cells.Add(new() {
            Point = new(0, 1),
            ValueGenerator = (col, row) => "建立日期:",
            CellStyleGenerator = (col, row) => SpreadsheetManager.DefaultCellStyles.CellStyle
        });

        cells.Add(new() {
            Point = new(1, 1),
            ValueGenerator = (col, row) => createdDate.ToString("yyyy-MM-dd HH:mm:ss"),
            CellStyleGenerator = (col, row) => SpreadsheetManager.DefaultCellStyles.CellStyle
        });

        rowHeights[1] = null;

        // 空白列
        rowHeights[2] = 10;

        return new TemplateContext(cells, 3, rowHeights);
    }
}
