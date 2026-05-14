using System.Text;
using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Sample;
using Json.Schema;

namespace CloudyWing.SpreadsheetExporter.Renderer.ClosedXML.Tests;

[TestFixture]
internal sealed class JsonSchemaValidationTests {
    private static readonly string RepositoryRoot = FindRepositoryRoot();
    private static readonly string SchemaPath = Path.Combine(
        RepositoryRoot,
        "docs",
        "schemas",
        "spreadsheet-exporter.schema.json"
    );
    private static readonly EvaluationOptions EvaluationOptions = new() {
        OutputFormat = OutputFormat.List
    };
    private static readonly Lazy<JsonSchema> Schema = new(LoadSchemaCore);

    [Test]
    public void SpreadsheetExporterSchema_SampleJsonTemplate_ShouldBeValid() {
        JsonSchema schema = LoadSchema();
        using JsonDocument template = JsonDocument.Parse(SampleWorkbookFactory.CreateJsonTemplate());

        EvaluationResults results = schema.Evaluate(template.RootElement, EvaluationOptions);

        Assert.That(results.IsValid, Is.True, CreateFailureMessage(results));
    }

    [Test]
    public void SpreadsheetExporterSchema_DocumentWorkbookJsonBlocks_ShouldBeValid() {
        JsonSchema schema = LoadSchema();
        string documentPath = Path.Combine(RepositoryRoot, "docs", "articles", "spreadsheet-document.md");
        IReadOnlyList<string> workbookJsonBlocks = ExtractJsonBlocks(File.ReadAllText(documentPath))
            .Where(IsWorkbookJson)
            .ToArray();

        Assert.That(workbookJsonBlocks, Has.Count.GreaterThanOrEqualTo(2));
        foreach (string json in workbookJsonBlocks) {
            using JsonDocument template = JsonDocument.Parse(json);
            EvaluationResults results = schema.Evaluate(template.RootElement, EvaluationOptions);

            Assert.That(results.IsValid, Is.True, CreateFailureMessage(results));
        }
    }

    [TestCase("""
        [
          {
            "FreezePanes": { "Row": 1 },
            "Templates": []
          }
        ]
        """)]
    [TestCase("""
        [
          {
            "Templates": [
              {
                "Type": "DataTable",
                "Columns": [
                  {}
                ]
              }
            ]
          }
        ]
        """)]
    [TestCase("""
        [
          {
            "Templates": [
              {
                "Type": "Grid",
                "Rows": [
                  {
                    "Cells": [
                      { "Value": "A", "Formula": "A1" }
                    ]
                  }
                ]
              }
            ]
          }
        ]
        """)]
    public void SpreadsheetExporterSchema_InvalidJsonTemplate_ShouldFail(string json) {
        JsonSchema schema = LoadSchema();
        using JsonDocument template = JsonDocument.Parse(json);

        EvaluationResults results = schema.Evaluate(template.RootElement, EvaluationOptions);

        Assert.That(results.IsValid, Is.False);
    }

    [Test]
    public void SpreadsheetExporterSchema_CustomTemplateType_ShouldBeValid() {
        JsonSchema schema = LoadSchema();
        using JsonDocument template = JsonDocument.Parse("""
            [
              {
                "Templates": [
                  {
                    "Type": "ReportHeader",
                    "Title": "Q1 Sales Report",
                    "ColumnSpan": 4
                  }
                ]
              }
            ]
            """);

        EvaluationResults results = schema.Evaluate(template.RootElement, EvaluationOptions);

        Assert.That(results.IsValid, Is.True, CreateFailureMessage(results));
    }

    private static JsonSchema LoadSchema() {
        return Schema.Value;
    }

    private static JsonSchema LoadSchemaCore() {
        using JsonDocument schemaDocument = JsonDocument.Parse(File.ReadAllText(SchemaPath));
        return JsonSchema.Build(schemaDocument.RootElement.Clone());
    }

    private static IReadOnlyList<string> ExtractJsonBlocks(string markdown) {
        List<string> blocks = [];
        StringBuilder? currentBlock = null;

        foreach (string line in markdown.Replace("\r\n", "\n", StringComparison.Ordinal).Split('\n')) {
            if (currentBlock is null) {
                if (line.Trim().Equals("```json", StringComparison.OrdinalIgnoreCase)) {
                    currentBlock = new StringBuilder();
                }

                continue;
            }

            if (line.Trim().Equals("```", StringComparison.Ordinal)) {
                blocks.Add(currentBlock.ToString());
                currentBlock = null;
                continue;
            }

            currentBlock.AppendLine(line);
        }

        return blocks.AsReadOnly();
    }

    private static bool IsWorkbookJson(string json) {
        using JsonDocument document = JsonDocument.Parse(json);
        return document.RootElement.ValueKind == JsonValueKind.Array;
    }

    private static string CreateFailureMessage(EvaluationResults results) {
        List<string> messages = [];
        CollectFailureMessages(results, messages);

        return messages.Count == 0
            ? "JSON Schema validation failed without detailed errors."
            : string.Join(Environment.NewLine, messages);
    }

    private static void CollectFailureMessages(EvaluationResults results, List<string> messages) {
        if (!results.IsValid && results.Errors is not null) {
            foreach (KeyValuePair<string, string> error in results.Errors) {
                messages.Add($"{results.InstanceLocation}: {error.Key} - {error.Value}");
            }
        }

        if (results.Details is null) {
            return;
        }

        foreach (EvaluationResults detail in results.Details) {
            CollectFailureMessages(detail, messages);
        }
    }

    private static string FindRepositoryRoot() {
        DirectoryInfo? directory = new(TestContext.CurrentContext.TestDirectory);
        while (directory is not null) {
            string solutionPath = Path.Combine(directory.FullName, "SpreadsheetExporter.slnx");
            if (File.Exists(solutionPath)) {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Cannot find SpreadsheetExporter.slnx from the test directory.");
    }
}
