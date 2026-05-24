using System.Text.Json;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.Json {
    [TestFixture]
    internal class JsonSchemaTests {
        private static readonly string RepositoryRoot = FindRepositoryRoot();
        private static readonly string SchemaPath = Path.Combine(
            RepositoryRoot,
            "docs",
            "schemas",
            "spreadsheet-exporter.schema.json"
        );

        [Test]
        public void SpreadsheetExporterSchema_ShouldBeValidJson() {
            using FileStream stream = File.OpenRead(SchemaPath);
            using JsonDocument document = JsonDocument.Parse(stream);

            using (Assert.EnterMultipleScope()) {
                Assert.That(document.RootElement.GetProperty("$schema").GetString(), Is.EqualTo(
                    "https://json-schema.org/draft/2020-12/schema"
                ));
                Assert.That(document.RootElement.GetProperty("type").GetString(), Is.EqualTo("array"));
            }
        }

        [Test]
        public void SpreadsheetExporterSchema_LocalReferences_ShouldResolve() {
            using FileStream stream = File.OpenRead(SchemaPath);
            using JsonDocument document = JsonDocument.Parse(stream);
            List<string> references = [];

            CollectReferences(document.RootElement, references);

            Assert.That(references, Is.Not.Empty);
            foreach (string reference in references) {
                Assert.That(ResolvePointer(document.RootElement, reference), Is.True, reference);
            }
        }

        [Test]
        public void SpreadsheetExporterSchema_ShouldDeclareBuiltInTemplateDefinitions() {
            using FileStream stream = File.OpenRead(SchemaPath);
            using JsonDocument document = JsonDocument.Parse(stream);
            JsonElement defs = document.RootElement.GetProperty("$defs");
            string[] templateReferences = defs
                .GetProperty("template")
                .GetProperty("anyOf")
                .EnumerateArray()
                .Select(x => x.GetProperty("$ref").GetString()!)
                .ToArray();

            using (Assert.EnterMultipleScope()) {
                Assert.That(templateReferences, Does.Contain("#/$defs/gridTemplate"));
                Assert.That(templateReferences, Does.Contain("#/$defs/recordSetTemplate"));
                Assert.That(templateReferences, Does.Contain("#/$defs/dataTableTemplate"));
                Assert.That(templateReferences, Does.Contain("#/$defs/mergedTemplate"));
            }
        }

        [Test]
        public void DocfxConfig_ShouldPublishSchemaResource() {
            string docfxConfigPath = Path.Combine(RepositoryRoot, "docs", "docfx.json");
            using FileStream stream = File.OpenRead(docfxConfigPath);
            using JsonDocument document = JsonDocument.Parse(stream);
            JsonElement resources = document.RootElement.GetProperty("build").GetProperty("resource");

            bool hasSchemaResource = resources
                .EnumerateArray()
                .SelectMany(x => x.GetProperty("files").EnumerateArray())
                .Select(x => x.GetString())
                .Contains("schemas/**");

            Assert.That(hasSchemaResource, Is.True);
        }

        private static void CollectReferences(JsonElement element, List<string> references) {
            if (element.ValueKind == JsonValueKind.Object) {
                foreach (JsonProperty property in element.EnumerateObject()) {
                    if (property.Name == "$ref" && property.Value.ValueKind == JsonValueKind.String) {
                        references.Add(property.Value.GetString()!);
                    }

                    CollectReferences(property.Value, references);
                }

                return;
            }

            if (element.ValueKind == JsonValueKind.Array) {
                foreach (JsonElement item in element.EnumerateArray()) {
                    CollectReferences(item, references);
                }
            }
        }

        private static bool ResolvePointer(JsonElement root, string reference) {
            if (!reference.StartsWith("#/", StringComparison.Ordinal)) {
                return false;
            }

            JsonElement current = root;
            string[] segments = reference[2..]
                .Split('/')
                .Select(DecodePointerSegment)
                .ToArray();

            foreach (string segment in segments) {
                if (current.ValueKind == JsonValueKind.Object) {
                    if (!current.TryGetProperty(segment, out current)) {
                        return false;
                    }

                    continue;
                }

                if (current.ValueKind == JsonValueKind.Array
                    && int.TryParse(segment, out int index)
                    && index >= 0
                    && index < current.GetArrayLength()
                ) {
                    current = current[index];
                    continue;
                }

                return false;
            }

            return true;
        }

        private static string DecodePointerSegment(string value) {
            return value.Replace("~1", "/", StringComparison.Ordinal)
                .Replace("~0", "~", StringComparison.Ordinal);
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
}
