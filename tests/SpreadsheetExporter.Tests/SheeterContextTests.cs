using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;

namespace CloudyWing.SpreadsheetExporter.Tests {
    [TestFixture]
    internal class SheeterContextTests {
        [Test]
        public void Constructor_ShouldInitializeProperties() {
            Sheeter sheeter = new("Sheet1") {
                Password = "password"
            };
            sheeter.PageSettings.PaperSize = PaperSize.A4;
            sheeter.SetColumnWidth(1, 50D);
            sheeter.SetColumnWidth(2, 75D);

            ITemplate template = Substitute.For<ITemplate>();
            template.GetContext().Returns(new TemplateContext(Enumerable.Empty<Cell>(), 0, new Dictionary<int, double?>()));
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            Assert.That(sut.SheetName, Is.EqualTo(sheeter.SheetName));
            Assert.That(sut.Cells, Is.Empty);
            Assert.That(sut.ColumnWidths, Is.EqualTo(sheeter.ColumnWidths));
            Assert.That(sut.Password, Is.EqualTo(sheeter.Password));
            Assert.That(sut.IsProtected, Is.True);
            Assert.That(sut.PageSettings, Is.EqualTo(sheeter.PageSettings));
        }

        [Test]
        public void Constructor_TemplateHasFreezePanes_ShouldCopyFreezePanesToSheeterContext() {
            Sheeter sheeter = new("Sheet1");
            Point freezePoint = new(1, 2);

            ITemplate template = Substitute.For<ITemplate>();
            TemplateContext templateContext = new(
                Enumerable.Empty<Cell>(),
                0,
                new Dictionary<int, double?>()
            ) {
                FreezePanes = freezePoint
            };
            template.GetContext().Returns(templateContext);
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            Assert.That(sut.FreezePanes, Is.EqualTo(freezePoint));
        }

        [Test]
        public void Constructor_TemplateHasNoFreezePanes_ShouldHaveNullFreezePanes() {
            Sheeter sheeter = new("Sheet1");

            ITemplate template = Substitute.For<ITemplate>();
            template.GetContext().Returns(new TemplateContext(
                Enumerable.Empty<Cell>(),
                0,
                new Dictionary<int, double?>()
            ));
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            Assert.That(sut.FreezePanes, Is.Null);
        }

        [Test]
        public void Constructor_TemplateHasAutoFilterEnabled_ShouldCopyAutoFilterEnabledToSheeterContext() {
            Sheeter sheeter = new("Sheet1");

            ITemplate template = Substitute.For<ITemplate>();
            TemplateContext templateContext = new(
                Enumerable.Empty<Cell>(),
                0,
                new Dictionary<int, double?>()
            ) {
                IsAutoFilterEnabled = true
            };
            template.GetContext().Returns(templateContext);
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            Assert.That(sut.IsAutoFilterEnabled, Is.True);
        }

        [Test]
        public void Constructor_TemplateHasAutoFilterDisabled_ShouldHaveAutoFilterDisabled() {
            Sheeter sheeter = new("Sheet1");

            ITemplate template = Substitute.For<ITemplate>();
            template.GetContext().Returns(new TemplateContext(
                Enumerable.Empty<Cell>(),
                0,
                new Dictionary<int, double?>()
            ));
            sheeter.AddTemplate(template);

            SheeterContext sut = new(sheeter);

            Assert.That(sut.IsAutoFilterEnabled, Is.False);
        }
    }
}
