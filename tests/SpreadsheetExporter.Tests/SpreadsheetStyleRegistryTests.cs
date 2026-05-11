namespace CloudyWing.SpreadsheetExporter.Tests;

[TestFixture]
internal class SpreadsheetStyleRegistryTests {
    [Test]
    public void Set_ShouldRegisterStyleByName() {
        CellStyle style = new(Font: new CellFont("Arial", 12));
        SpreadsheetStyleRegistry sut = new();

        sut.Set("Header", style);

        using (Assert.EnterMultipleScope()) {
            Assert.That(sut.TryGet("Header", out CellStyle actual), Is.True);
            Assert.That(actual, Is.EqualTo(style));
        }
    }

    [Test]
    public void Import_ShouldCopyStylesFromAnotherRegistry() {
        CellStyle style = new(HasBorder: true);
        SpreadsheetStyleRegistry source = new();
        source.Set("Bordered", style);
        SpreadsheetStyleRegistry sut = new();

        sut.Import(source);

        using (Assert.EnterMultipleScope()) {
            Assert.That(sut.TryGet("Bordered", out CellStyle actual), Is.True);
            Assert.That(actual, Is.EqualTo(style));
        }
    }

    [Test]
    public void Clear_ShouldRemoveRegisteredStyles() {
        SpreadsheetStyleRegistry sut = new();
        sut.Set("Header", new CellStyle());

        sut.Clear();

        Assert.That(sut.TryGet("Header", out _), Is.False);
    }
}
