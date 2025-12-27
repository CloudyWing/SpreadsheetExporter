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

            Assert.That(contentType, Is.EqualTo(expected));
        }

        [Test]
        public void FileNameExtension_ShouldReturnCorrectFileNameExtension() {
            string expected = FakeFileNameExtension;
            string fileNameExtension = exporter!.FileNameExtension;


            Assert.That(fileNameExtension, Is.EqualTo(expected));
        }

        [Test]
        public void Password_ShouldSetCorrectPassword() {
            string expected = "123456";
            exporter!.Password = expected;


            Assert.That(exporter!.Password, Is.EqualTo(expected));
        }

        [Test]
        public void HasPassword_WhenPasswordIsSet_ShouldReturnTrue() {
            exporter!.Password = "123456";

            bool hasPassword = exporter!.HasPassword;

            Assert.That(hasPassword, Is.True);
        }

        [Test]
        public void HasPassword_WhenPasswordIsNotSet_ShouldReturnFalse() {
            exporter!.Password = "";
            bool hasPassword = exporter!.HasPassword;

            Assert.That(hasPassword, Is.False);
        }

        [Test]
        public void DefaultBasicSheetName_ShouldSetCorrectDefaultBasicSheetName() {
            string expected = "工作表1";
            exporter!.DefaultBasicSheetName = expected;

            Assert.That(exporter!.DefaultBasicSheetName, Is.EqualTo(expected));
        }

        [Test]
        public void LastSheeter_ShouldReturnCorrectLastSheeter() {
            exporter!.CreateSheeter("Sheet1");
            Sheeter sheeter = exporter!.CreateSheeter("Sheet2");

            Sheeter lastSheeter = exporter!.LastSheeter;

            Assert.That(lastSheeter, Is.EqualTo(sheeter));
        }

        [Test]
        public void CreateSheeter_WhenSheetNameIsProvided_ShouldCreateNewSheeter() {
            string sheetName = "Sheet1";
            double rowHeight = 20D;

            Sheeter newSheeter = exporter!.CreateSheeter(sheetName, rowHeight);

            Assert.That(newSheeter.SheetName, Is.EqualTo(sheetName));
            Assert.That(newSheeter.DefaultRowHeight, Is.EqualTo(rowHeight));
        }

        [Test]
        public void GetSheeter_ShouldReturnCorrectSheeter() {
            exporter!.CreateSheeter("Sheet1");
            Sheeter sheeter = exporter!.CreateSheeter("Sheet2");

            Sheeter returnedSheeter = exporter!.GetSheeter(1);

            Assert.That(returnedSheeter, Is.EqualTo(sheeter));
        }

        [Test]
        public void GetSheeter_WhenIndexIsNegative_ShouldThrowArgumentOutOfRangeException() {
            exporter!.CreateSheeter("Sheet1");

            Action action = () => exporter!.GetSheeter(-1);

            ArgumentOutOfRangeException? ex = Assert.Throws<ArgumentOutOfRangeException>(() => action());
            Assert.That(ex!.ParamName, Is.EqualTo("index"));
            Assert.That(ex.ActualValue, Is.EqualTo(-1));
        }

        [Test]
        public void GetSheeter_WhenIndexIsOutOfRange_ShouldThrowArgumentOutOfRangeException() {
            exporter!.CreateSheeter("Sheet1");

            Action action = () => exporter!.GetSheeter(1);

            ArgumentOutOfRangeException? ex = Assert.Throws<ArgumentOutOfRangeException>(() => action());
            Assert.That(ex!.ParamName, Is.EqualTo("index"));
            Assert.That(ex.ActualValue, Is.EqualTo(1));
            Assert.That(ex.Message, Does.Contain("Index must be between 0 and 0"));
        }

        [Test]
        public void TryGetSheeter_WhenIndexIsValid_ShouldReturnTrueAndSheeter() {
            exporter!.CreateSheeter("Sheet1");
            Sheeter expectedSheeter = exporter!.CreateSheeter("Sheet2");

            bool result = exporter!.TryGetSheeter(1, out Sheeter sheeter);

            Assert.That(result, Is.True);
            Assert.That(sheeter, Is.EqualTo(expectedSheeter));
        }

        [Test]
        public void TryGetSheeter_WhenIndexIsNegative_ShouldReturnFalseAndNullSheeter() {
            exporter!.CreateSheeter("Sheet1");

            bool result = exporter!.TryGetSheeter(-1, out Sheeter sheeter);

            Assert.That(result, Is.False);
            Assert.That(sheeter, Is.Null);
        }

        [Test]
        public void TryGetSheeter_WhenIndexIsOutOfRange_ShouldReturnFalseAndNullSheeter() {
            exporter!.CreateSheeter("Sheet1");

            bool result = exporter!.TryGetSheeter(1, out Sheeter sheeter);

            Assert.That(result, Is.False);
            Assert.That(sheeter, Is.Null);
        }

        [Test]
        public void TryGetSheeter_WhenNoSheetersExist_ShouldReturnFalseAndNullSheeter() {
            bool result = exporter!.TryGetSheeter(0, out Sheeter sheeter);

            Assert.That(result, Is.False);
            Assert.That(sheeter, Is.Null);
        }

        [Test]
        public void Export_WhenThereAreNoSheeters_ShouldThrowSheeterNotFoundException() {
            Action action = () => exporter!.Export();
            Assert.Throws<SheeterNotFoundException>(() => action());
        }

        [Test]
        public void Export_BeforeExecutingExport_ShouldInvokeSpreadsheetExportingEvent() {
            exporter!.CreateSheeter();
            bool spreadsheetExportingEventInvoked = false;
            exporter!.SpreadsheetExportingEvent += (sender, args) => spreadsheetExportingEventInvoked = true;

            exporter!.Export();

            Assert.That(spreadsheetExportingEventInvoked, Is.True);
        }

        [Test]
        public void Export_AfterExecutingExport_ShouldInvokeSpreadsheetExportedEvent() {
            exporter!.CreateSheeter();
            bool spreadsheetExportedEventInvoked = false;
            exporter!.SpreadsheetExportedEvent += (sender, args) => spreadsheetExportedEventInvoked = true;

            exporter!.Export();

            Assert.That(spreadsheetExportedEventInvoked, Is.True);
        }

        [Test]
        public void ExportFile_ShouldCreateNewFile_WhenFileModeIsCreate() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create;

            exporter!.ExportFile(FilePath, fileMode);

            Assert.That(File.Exists(FilePath), Is.True);
        }

        [Test]
        public void ExportFile_WhenFileModeIsCreateNew_ShouldCreateNewFile() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.CreateNew;

            exporter!.ExportFile(FilePath, fileMode);

            Assert.That(File.Exists(FilePath), Is.True);
        }

        [Test]
        public void ExportFile_WhenFileModeIsCreateNewAndFileAlreadyExists_ShouldThrowIOException() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.CreateNew;

            exporter!.ExportFile(FilePath, SpreadsheetFileMode.Create);

            Action action = () => exporter!.ExportFile(FilePath, fileMode);

            Assert.Throws<IOException>(() => action());
        }

        [Test]
        public void ExportFile_ShouldWriteExportedBytesToFile() {
            exporter!.CreateSheeter();
            SpreadsheetFileMode fileMode = SpreadsheetFileMode.Create;

            exporter!.ExportFile(FilePath, fileMode);

            using FileStream fileStream = File.OpenRead(FilePath);
            byte[] fileBytes = new byte[fileStream.Length];
            fileStream.Read(fileBytes, 0, fileBytes.Length);
            Assert.That(fileBytes, Is.EqualTo(fakeExportResult));
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
