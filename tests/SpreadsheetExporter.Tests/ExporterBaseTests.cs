using CloudyWing.SpreadsheetExporter;
using CloudyWing.SpreadsheetExporter.Exceptions;

namespace CloudyWing.Spreadsheetexporter.Tests {
    internal class ExporterBaseTests {
        private const string FilePath = "test.txt";
        private const string FakeContentType = "application/ms-excel";
        private const string FakeFileNameExtension = ".xlsx";
        private static readonly byte[] fakeExportResult = new byte[] { 0x48, 0x65, 0x6C, 0x6C, 0x6F };
        private ExporterBase? exporter;

        [SetUp]
        public void SetUp() {
            exporter = new FakeExporter();
        }

        [Test]
        public void ContentType_ShouldReturnCorrectContentType() {
            string expected = FakeContentType;

            string contentType = exporter!.ContentType;

            contentType.Should().Be(expected);
        }

        [Test]
        public void FileNameExtension_ShouldReturnCorrectFileNameExtension() {
            string expected = FakeFileNameExtension;
            string fileNameExtension = exporter!.FileNameExtension;


            fileNameExtension.Should().Be(expected);
        }

        [Test]
        public void Password_ShouldSetCorrectPassword() {
            string expected = "123456";
            exporter!.Password = expected;


            exporter!.Password.Should().Be(expected);
        }

        [Test]
        public void HasPassword_WhenPasswordIsSet_ShouldReturnTrue() {
            exporter!.Password = "123456";

            bool hasPassword = exporter!.HasPassword;

            hasPassword.Should().BeTrue();
        }

        [Test]
        public void HasPassword_WhenPasswordIsNotSet_ShouldReturnFalse() {
            exporter!.Password = "";
            bool hasPassword = exporter!.HasPassword;

            hasPassword.Should().BeFalse();
        }

        [Test]
        public void DefaultBasicSheetName_ShouldSetCorrectDefaultBasicSheetName() {
            string expected = "工作表1";
            exporter!.DefaultBasicSheetName = expected;

            exporter!.DefaultBasicSheetName.Should().Be(expected);
        }

        [Test]
        public void LastSheeter_ShouldReturnCorrectLastSheeter() {
            exporter!.CreateSheeter("Sheet1");
            Sheeter sheeter = exporter!.CreateSheeter("Sheet2");

            Sheeter lastSheeter = exporter!.LastSheeter;

            lastSheeter.Should().Be(sheeter);
        }

        [Test]
        public void CreateSheeter_WhenSheetNameIsProvided_ShouldCreateNewSheeter() {
            string sheetName = "Sheet1";
            double rowHeight = 20D;

            Sheeter newSheeter = exporter!.CreateSheeter(sheetName, rowHeight);

            newSheeter.SheetName.Should().Be(sheetName);
            newSheeter.DefaultRowHeight.Should().Be(rowHeight);
        }

        [Test]
        public void GetSheeter_ShouldReturnCorrectSheeter() {
            exporter!.CreateSheeter("Sheet1");
            Sheeter sheeter = exporter!.CreateSheeter("Sheet2");

            Sheeter returnedSheeter = exporter!.GetSheeter(1);

            returnedSheeter.Should().Be(sheeter);
        }

        [Test]
        public void Export_WhenThereAreNoSheeters_ShouldThrowSheeterNotFoundException() {
            Action action = () => exporter!.Export();
            action.Should().Throw<SheeterNotFoundException>();
        }

        [Test]
        public void Export_BeforeExecutingExport_ShouldInvokeSpreadsheetExportingEvent() {
            exporter!.CreateSheeter();
            bool spreadsheetExportingEventInvoked = false;
            exporter!.SpreadsheetExportingEvent += (sender, args) => spreadsheetExportingEventInvoked = true;

            exporter!.Export();

            spreadsheetExportingEventInvoked.Should().BeTrue();
        }

        [Test]
        public void Export_AfterExecutingExport_ShouldInvokeSpreadsheetExportedEvent() {
            exporter!.CreateSheeter();
            bool spreadsheetExportedEventInvoked = false;
            exporter!.SpreadsheetExportedEvent += (sender, args) => spreadsheetExportedEventInvoked = true;

            exporter!.Export();

            spreadsheetExportedEventInvoked.Should().BeTrue();
        }

        [Test]
        public void ExportFile_ShouldCreateNewFile_WhenFileModeIsCreate() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create;

            exporter!.ExportFile(FilePath, fileMode);

            File.Exists(FilePath).Should().BeTrue();
        }

        [Test]
        public void ExportFile_WhenFileModeIsCreateNew_ShouldCreateNewFile() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.CreateNew;

            exporter!.ExportFile(FilePath, fileMode);

            File.Exists(FilePath).Should().BeTrue();
        }

        [Test]
        public void ExportFile_WhenFileModeIsCreateNewAndFileAlreadyExists_ShouldThrowIOException() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.CreateNew;

            exporter!.ExportFile(FilePath, SpreadsheetFileMode.Create);

            Action action = () => exporter!.ExportFile(FilePath, fileMode);

            action.Should().Throw<IOException>();
        }

        [Test]
        public void ExportFile_ShouldWriteExportedBytesToFile() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create;

            exporter!.ExportFile(FilePath, fileMode);

            using FileStream fileStream = File.OpenRead(FilePath);
            byte[] fileBytes = new byte[fileStream.Length];
            fileStream.Read(fileBytes, 0, fileBytes.Length);
            fileBytes.Should().BeEquivalentTo(fakeExportResult);
        }

        [TearDown]
        public void TearDown() {
            if (File.Exists(FilePath)) {
                File.Delete(FilePath);
            }
        }

        private class FakeExporter : ExporterBase {
            public override string ContentType => FakeContentType;

            public override string FileNameExtension => FakeFileNameExtension;

            protected override byte[] ExecuteExport(IEnumerable<SheeterContext> contexts) {
                return fakeExportResult;
            }
        }
    }
}
