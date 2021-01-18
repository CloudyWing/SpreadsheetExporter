using System.Drawing;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using FluentAssertions;
using NUnit.Framework;

namespace CloudyWing.SpreadsheetExporter.Config.Tests {
    [TestFixture]
    public class SpreadsheetConfigurationTests {
        public const string JsonFile = "Spreadsheet.json";

        private ConfigSettings CreateConfigSettingsByJsonFile() {
            return new ConfigSettings {
                Cell = new CellSettings {
                    HorizontalAlignment = HorizontalAlignment.Center,
                    VerticalAlignment = VerticalAlignment.Middle,
                    HasBorder = false,
                    WrapText = false,
                    Font = new FontSettings {
                        FontName = "新細明體",
                        FontSize = 10,
                        IsBold = false,
                        IsItalic = false,
                        HasUnderline = false,
                        IsStrikeout = false
                    }
                }
            };
        }

        private async Task WriteJsonFileAsync(ConfigSettings configSettings) {
            object settings = new {
                Spreadsheet = configSettings
            };

            await File.WriteAllTextAsync(
                Path.Combine(TestContext.CurrentContext.TestDirectory, JsonFile),
                JsonSerializer.Serialize(settings)
            );
        }

        private void DeleteJsonFile() {
            File.Delete(Path.Combine(TestContext.CurrentContext.TestDirectory, JsonFile));
        }

        private void CreateTestCellStyle(
            CellStyle cellStyle, out CellStyle listHeaderStyle,
            out CellStyle listTextStyle, out CellStyle listNumberStyle, out CellStyle listDateTimeStyle
        ) {
            CellFont listHeaderFont = cellStyle.Font
                    .CloneAndSetStyle(cellStyle.Font.Style | FontStyles.IsBold);

            listHeaderStyle = cellStyle
                    .CloneAndSetFont(listHeaderFont)
                    .CloneAndSetHorizontalAlignment(HorizontalAlignment.Center)
                    .CloneAndSetBorder(true)
                    .CloneAndSetFont(listHeaderFont);

            listTextStyle = cellStyle
                .CloneAndSetBorder(true)
                .CloneAndSetHorizontalAlignment(HorizontalAlignment.Left);

            listNumberStyle = cellStyle
                .CloneAndSetBorder(true)
                .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);

            listDateTimeStyle = cellStyle
                .CloneAndSetBorder(true)
                .CloneAndSetHorizontalAlignment(HorizontalAlignment.Right);
        }

        [Test]
        public async Task Constructor_UseJsonFile_ReturnTrue() {
            ConfigSettings configSettings = CreateConfigSettingsByJsonFile();

            await WriteJsonFileAsync(configSettings);

            SpreadsheetConfiguration actual = new SpreadsheetConfiguration(
                TestContext.CurrentContext.TestDirectory, JsonFile
            );

            CellStyle cellStyle = new CellStyle(font: new CellFont("新細明體", 10, Color.Black));

            CreateTestCellStyle(cellStyle, out CellStyle listHeaderStyle,
                out CellStyle listTextStyle, out CellStyle listNumberStyle, out CellStyle listDateTimeStyle
            );

            actual.CellStyle.Should().Be(cellStyle);
            actual.ListHeaderStyle.Should().Be(listHeaderStyle);
            actual.ListTextStyle.Should().Be(listTextStyle);
            actual.ListNumberStyle.Should().Be(listNumberStyle);
            actual.ListDateTimeStyle.Should().Be(listDateTimeStyle);

            DeleteJsonFile();
        }

        [Test]
        public async Task Constructor_ChangeJsonFile_ReturnTrue() {
            ConfigSettings configSettings = CreateConfigSettingsByJsonFile();
            await WriteJsonFileAsync(configSettings);

            SpreadsheetConfiguration actual = new SpreadsheetConfiguration(
                TestContext.CurrentContext.TestDirectory, JsonFile
            );

            configSettings.Cell.Font.FontName = "標楷體";
            await WriteJsonFileAsync(configSettings);

            // 要等ChangeToken.OnChange()觸發完畢
            await Task.Delay(5000);
            actual.CellStyle.Font.Name.Should().Be("標楷體");

            DeleteJsonFile();
        }

        [Test]
        public void Constructor_UseAction_ReturnTrue() {
            CellFont cellFont = new CellFont("新細明體", 10, Color.Black, FontStyles.None);
            CellStyle cellStyle = new CellStyle(HorizontalAlignment.Center, VerticalAlignment.Middle, false, false, Color.Empty, cellFont);

            CreateTestCellStyle(cellStyle, out CellStyle listHeaderStyle,
                out CellStyle listTextStyle, out CellStyle listNumberStyle, out CellStyle listDateTimeStyle
            );

            SpreadsheetConfiguration actual = new SpreadsheetConfiguration(x => {
                x.CellStyle = cellStyle;
                x.ListHeaderStyle = listHeaderStyle;
                x.ListTextStyle = listTextStyle;
                x.ListNumberStyle = listNumberStyle;
                x.ListDateTimeStyle = listDateTimeStyle;
            });

            actual.CellStyle.Should().Be(cellStyle);
            actual.ListHeaderStyle.Should().Be(listHeaderStyle);
            actual.ListTextStyle.Should().Be(listTextStyle);
            actual.ListNumberStyle.Should().Be(listNumberStyle);
            actual.ListDateTimeStyle.Should().Be(listDateTimeStyle);
        }
    }
}