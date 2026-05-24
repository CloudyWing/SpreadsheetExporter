using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Renderer.ClosedXML;
using CloudyWing.SpreadsheetExporter.Sample;

SpreadsheetManager.SetRenderer(static () => new ExcelRenderer());

string outputDirectory = Path.Combine(AppContext.BaseDirectory, "artifacts");
Directory.CreateDirectory(outputDirectory);

SpreadsheetDocument fluentDocument = SampleWorkbookFactory.CreateFluentDocument();
string fluentPath = Path.Combine(outputDirectory, $"sales-report{fluentDocument.FileNameExtension}");
fluentDocument.ExportFile(fluentPath);

SpreadsheetDocument jsonDocument = SampleWorkbookFactory.CreateJsonDocument();
string jsonPath = Path.Combine(outputDirectory, $"sales-report-json{jsonDocument.FileNameExtension}");
jsonDocument.ExportFile(jsonPath);

Console.WriteLine("SpreadsheetExporter Sample (ClosedXML)");
Console.WriteLine($"Fluent API workbook : {fluentPath}");
Console.WriteLine($"JSON template output: {jsonPath}");
