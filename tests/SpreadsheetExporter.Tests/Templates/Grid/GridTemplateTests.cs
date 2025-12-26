using System.Drawing;
using CloudyWing.SpreadsheetExporter.Templates;
using CloudyWing.SpreadsheetExporter.Templates.Grid;

namespace CloudyWing.SpreadsheetExporter.Tests.Templates.Grid {
    [TestFixture]
    internal class GridTemplateTests {
        [Test]
        public void CreateRow_ShouldAddRowToTemplate() {
            GridTemplate template = new();
            template.CreateRow();

            Assert.That(template.RowSpan, Is.EqualTo(1));
        }

        [Test]
        public void CreateRow_SetRowHeight_ShouldAddRowWithSpecifiedHeightToTemplate() {
            GridTemplate template = new();
            template.CreateRow(10);

            Assert.That(template.GetContext().RowHeights[0], Is.EqualTo(10));
        }

        [Test]
        public void CreateCell_ShouldAddCellToLastRow() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1");

            Assert.That(template.GetContext().Cells.Count, Is.EqualTo(1));
        }

        [Test]
        public void CreateCell_SetCellValue_ShouldAddCellWithSpecifiedValue() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1");

            Assert.That(template.GetContext().Cells.Single().GetValue(), Is.EqualTo("A1"));
        }

        [Test]
        public void CreateCell_SetColumnSpan_ShouldAddCellWithSpecifiedColumnSpan() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", 2);

            Assert.That(template.GetContext().Cells.Single().Size.Width, Is.EqualTo(2));
        }

        [Test]
        public void CreateCell_SetRowSpan_ShouldAddCellWithSpecifiedRowSpan() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", rowSpan: 2);

            Assert.That(template.GetContext().Cells.Single().Size.Height, Is.EqualTo(2));
        }

        [Test]
        public void CreateCell_SetCellStyle_ShouldAddCellWithSpecifiedStyle() {
            CellStyle cellStyle = new(Font: new CellFont("Arial", 10));

            GridTemplate template = new();
            template.CreateRow()
                .CreateCell("A1", cellStyle: cellStyle);

            Assert.That(template.GetContext().Cells.Single().GetCellStyle(), Is.EqualTo(cellStyle));
        }

        [Test]
        public void CreateCell_SetFormula_ShouldCreateCellWithFormula() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell((x, y) => "A1 + B1", 2, 2, SpreadsheetManager.DefaultCellStyles.GridCellStyle);

            TemplateContext context = template.GetContext();
            Cell cell = context.Cells.Single();
            Assert.That(cell.GetFormula(), Is.EqualTo("A1 + B1"));
            Assert.That(cell.GetCellStyle(), Is.EqualTo(SpreadsheetManager.DefaultCellStyles.GridCellStyle));
            Assert.That(cell.Size, Is.EqualTo(new Size(2, 2)));
        }

        [Test]
        public void GetContext_ShouldReturnExpectedTemplateContext() {
            GridTemplate template = new GridTemplate()
                .CreateRow(20)
                .CreateCell("A1", columnSpan: 2, rowSpan: 2)
                .CreateRow()
                .CreateCell("B1", columnSpan: 2, rowSpan: 2);

            TemplateContext context = template.GetContext();

            Assert.That(context.Cells.Count, Is.EqualTo(2));
            Assert.That(context.Cells[0].Point, Is.EqualTo(new Point(0, 0)));
            Assert.That(context.Cells[0].Size, Is.EqualTo(new Size(2, 2)));
            Assert.That(context.Cells[0].GetValue(), Is.EqualTo("A1"));
            Assert.That(context.Cells[1].Point, Is.EqualTo(new Point(2, 1)));
            Assert.That(context.Cells[1].Size, Is.EqualTo(new Size(2, 2)));
            Assert.That(context.Cells[1].GetValue(), Is.EqualTo("B1"));
            Assert.That(context.RowSpan, Is.EqualTo(3));
            Assert.That(context.RowHeights.Count, Is.EqualTo(2));
            Assert.That(context.RowHeights[0], Is.EqualTo(20));
            Assert.That(context.RowHeights[1], Is.EqualTo(null));
        }

        [Test]
        public void CreateCell_WithConfiguratorAndDataValidation_ShouldCreateCellWithDataValidation() {
            DataValidation validation = new() {
                ValidationType = DataValidationType.List,
                ListItems = new[] { "Option1", "Option2", "Option3" }
            };

            GridTemplate template = new();
            template.CreateRow()
                .CreateCell(cell => {
                    cell.ValueGenerator = (x, y) => "Option1";
                    cell.DataValidationGenerator = (x, y) => validation;
                });

            Cell resultCell = template.GetContext().Cells.Single();
            DataValidation? resultValidation = resultCell.GetDataValidation();

            Assert.That(resultValidation, Is.Not.Null);
            Assert.That(resultValidation.ValidationType, Is.EqualTo(DataValidationType.List));
            Assert.That(resultValidation.ListItems, Is.EqualTo(new[] { "Option1", "Option2", "Option3" }));
        }

        [Test]
        public void CreateCell_WithConfigurator_ShouldSupportListValidation() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell(cell => {
                    cell.ValueGenerator = (x, y) => "Red";
                    cell.DataValidationGenerator = (x, y) => new DataValidation {
                        ValidationType = DataValidationType.List,
                        ListItems = new[] { "Red", "Green", "Blue" },
                        IsDropdownShown = true,
                        ErrorTitle = "Invalid Color",
                        ErrorMessage = "Please select a color from the list"
                    };
                });

            DataValidation? validation = template.GetContext().Cells.Single().GetDataValidation();

            Assert.Multiple(() => {
                Assert.That(validation, Is.Not.Null);
                Assert.That(validation.ValidationType, Is.EqualTo(DataValidationType.List));
                Assert.That(validation.ListItems, Is.EquivalentTo(new[] { "Red", "Green", "Blue" }));
                Assert.That(validation.IsDropdownShown, Is.True);
                Assert.That(validation.ErrorTitle, Is.EqualTo("Invalid Color"));
                Assert.That(validation.ErrorMessage, Is.EqualTo("Please select a color from the list"));
            });
        }

        [Test]
        public void CreateCell_WithConfigurator_ShouldSupportIntegerValidation() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell(cell => {
                    cell.ValueGenerator = (x, y) => 50;
                    cell.DataValidationGenerator = (x, y) => new DataValidation {
                        ValidationType = DataValidationType.Integer,
                        Operator = DataValidationOperator.Between,
                        Value1 = 1,
                        Value2 = 100
                    };
                });

            DataValidation? validation = template.GetContext().Cells.Single().GetDataValidation();

            Assert.Multiple(() => {
                Assert.That(validation, Is.Not.Null);
                Assert.That(validation.ValidationType, Is.EqualTo(DataValidationType.Integer));
                Assert.That(validation.Operator, Is.EqualTo(DataValidationOperator.Between));
                Assert.That(validation.Value1, Is.EqualTo(1));
                Assert.That(validation.Value2, Is.EqualTo(100));
            });
        }

        [Test]
        public void CreateCell_WithConfigurator_ShouldSupportDateValidation() {
            DateTime minDate = new(2024, 1, 1);
            DateTime maxDate = new(2024, 12, 31);

            GridTemplate template = new();
            template.CreateRow()
                .CreateCell(cell => {
                    cell.ValueGenerator = (x, y) => DateTime.Now;
                    cell.DataValidationGenerator = (x, y) => new DataValidation {
                        ValidationType = DataValidationType.Date,
                        Operator = DataValidationOperator.Between,
                        Value1 = minDate,
                        Value2 = maxDate
                    };
                });

            DataValidation? validation = template.GetContext().Cells.Single().GetDataValidation();

            Assert.Multiple(() => {
                Assert.That(validation, Is.Not.Null);
                Assert.That(validation.ValidationType, Is.EqualTo(DataValidationType.Date));
                Assert.That(validation.Value1, Is.EqualTo(minDate));
                Assert.That(validation.Value2, Is.EqualTo(maxDate));
            });
        }

        [Test]
        public void CreateCell_WithConfigurator_ShouldSupportCustomValidation() {
            GridTemplate template = new();
            template.CreateRow()
                .CreateCell(cell => {
                    cell.ValueGenerator = (x, y) => "Test";
                    cell.DataValidationGenerator = (x, y) => new DataValidation {
                        ValidationType = DataValidationType.Custom,
                        Formula = "LEN(A1)<=10",
                        ErrorMessage = "Text must be 10 characters or less"
                    };
                });

            DataValidation? validation = template.GetContext().Cells.Single().GetDataValidation();

            Assert.Multiple(() => {
                Assert.That(validation, Is.Not.Null);
                Assert.That(validation.ValidationType, Is.EqualTo(DataValidationType.Custom));
                Assert.That(validation.Formula, Is.EqualTo("LEN(A1)<=10"));
                Assert.That(validation.ErrorMessage, Is.EqualTo("Text must be 10 characters or less"));
            });
        }
    }
}
