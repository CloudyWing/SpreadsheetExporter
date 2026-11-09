using System.Text.Json;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Grid;
using CloudyWing.SpreadsheetExporter.Templates.Json;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.Json;

[TestFixture]
internal class JsonTemplateRegistryTests {
    [Test]
    public void RegisterAndCreate_ConcurrentCalls_ShouldNotThrowAndKeepAllRegistrations() {
        string prefix = $"Concurrent-{Guid.NewGuid():N}-";
        using JsonDocument gridDocument = JsonDocument.Parse("""{ "Type": "Grid", "Rows": [] }""");
        JsonElement gridElement = gridDocument.RootElement.Clone();

        Parallel.For(0, 200, i => {
            JsonTemplateRegistry.Register(new StubParser(prefix + i));
            JsonTemplateRegistry.Create("Grid", gridElement);
        });

        Assert.That(
            Enumerable.Range(0, 200).Select(i => JsonTemplateRegistry.Create(prefix + i, gridElement)),
            Is.All.InstanceOf<GridTemplate>()
        );
    }

    private sealed class StubParser(string typeName) : ITemplateJsonParser {
        public string TypeName { get; } = typeName;

        public ISheetTemplate Parse(JsonElement element) {
            return new GridTemplate();
        }
    }
}
