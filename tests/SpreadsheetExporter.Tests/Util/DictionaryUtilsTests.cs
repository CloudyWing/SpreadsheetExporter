using CloudyWing.SpreadsheetExporter.Util;

namespace CloudyWing.SpreadsheetExporter.Tests.Util {
    [TestFixture]
    internal class DictionaryUtilsTests {
        private readonly TestRecord record = new() {
            Id = 1,
            Name = "Test",
            Address = new Address {
                Street = "Test Street",
                City = "Test City",
                Country = "Test Country"
            }
        };

        [Test]
        public void ConvertFrom_NestedPropertyLevelIs1_ShouldWithoutNestedProperties() {
            IDictionary<string, object?> result = DictionaryUtils.ConvertFrom(record, 1);

            Assert.Multiple(() => {
                Assert.That(result.Count, Is.EqualTo(3));
                Assert.That(result.ContainsKey("Id"), Is.True);
                Assert.That(result["Id"], Is.EqualTo(1));
                Assert.That(result.ContainsKey("Name"), Is.True);
                Assert.That(result["Name"], Is.EqualTo("Test"));
                Assert.That(result.ContainsKey("Address"), Is.True);
                Assert.That(result["Address"], Is.EqualTo(record.Address));
            });
        }

        [Test]
        public void ConvertFrom_NestedPropertyLevelIs2_ShouldIncludeNestedProperties() {
            IDictionary<string, object?> result = DictionaryUtils.ConvertFrom(record, 2);

            Assert.Multiple(() => {
                Assert.That(result.Count, Is.EqualTo(5));
                Assert.That(result.ContainsKey("Id"), Is.True);
                Assert.That(result["Id"], Is.EqualTo(1)); // 雖然在第一層，但基本型別直接回傳
                Assert.That(result.ContainsKey("Name"), Is.True);
                Assert.That(result["Name"], Is.EqualTo("Test")); // 雖然在第一層，但字串直接回傳
                Assert.That(result.ContainsKey("Address"), Is.False); // 因為抓到第二層，第一層的 Address 直接回傳
                Assert.That(result.ContainsKey("Address.Street"), Is.True);
                Assert.That(result["Address.Street"], Is.EqualTo("Test Street"));
                Assert.That(result.ContainsKey("Address.City"), Is.True);
                Assert.That(result["Address.City"], Is.EqualTo("Test City"));
                Assert.That(result.ContainsKey("Address.Country"), Is.True);
                Assert.That(result["Address.Country"], Is.EqualTo("Test Country"));
            });
        }

        [Test]
        public void ConvertFrom_TypeHasIndexer_ShouldSkipIndexer() {
            IndexedRecord indexedRecord = new() { Name = "Indexed" };

            IDictionary<string, object?> result = DictionaryUtils.ConvertFrom(indexedRecord, 1);

            Assert.Multiple(() => {
                Assert.That(result.Keys, Is.EquivalentTo(new[] { "Name" }));
                Assert.That(result["Name"], Is.EqualTo("Indexed"));
            });
        }

        [Test]
        public void ConvertFrom_PropertyHiddenByNew_ShouldUseMostDerivedProperty() {
            DerivedRecord derivedRecord = new() { Code = "Derived" };
            ((BaseRecord)derivedRecord).Code = 5;

            IDictionary<string, object?> result = DictionaryUtils.ConvertFrom(derivedRecord, 1);

            Assert.Multiple(() => {
                Assert.That(result, Has.Count.EqualTo(1));
                Assert.That(result["Code"], Is.EqualTo("Derived"));
            });
        }

        [TestCase("Name", "Test")]
        [TestCase("name", "Test")]
        [TestCase("Address.City", "Test City")]
        [TestCase("ADDRESS.CITY", "Test City")]
        public void TryGetPathValue_PathExists_ShouldReturnValue(string path, string expected) {
            bool isFound = DictionaryUtils.TryGetPathValue(record, path, out object? value, out _);

            Assert.Multiple(() => {
                Assert.That(isFound, Is.True);
                Assert.That(value, Is.EqualTo(expected));
            });
        }

        [Test]
        public void TryGetPathValue_IntermediateObjectIsNull_ShouldReturnNullValue() {
            TestRecord recordWithoutAddress = new() { Id = 2, Name = "NoAddress" };

            bool isFound = DictionaryUtils.TryGetPathValue(recordWithoutAddress, "Address.City", out object? value, out _);

            Assert.Multiple(() => {
                Assert.That(isFound, Is.True);
                Assert.That(value, Is.Null);
            });
        }

        [Test]
        public void TryGetPathValue_SourceIsNull_ShouldReturnNullValue() {
            bool isFound = DictionaryUtils.TryGetPathValue(null, "Name", out object? value, out _);

            Assert.Multiple(() => {
                Assert.That(isFound, Is.True);
                Assert.That(value, Is.Null);
            });
        }

        [Test]
        public void TryGetPathValue_PathDoesNotExist_ShouldReturnAvailablePathsAtFailedLevel() {
            bool isFound = DictionaryUtils.TryGetPathValue(
                record, "Address.Zip", out _, out IReadOnlyList<string> availablePaths
            );

            Assert.Multiple(() => {
                Assert.That(isFound, Is.False);
                Assert.That(
                    availablePaths,
                    Is.EquivalentTo(new[] { "Address.Street", "Address.City", "Address.Country" })
                );
            });
        }

        [Test]
        public void TryGetPathValue_TypeHasIndexer_ShouldResolveRegularProperty() {
            IndexedRecord indexedRecord = new() { Name = "Indexed" };

            bool isFound = DictionaryUtils.TryGetPathValue(indexedRecord, "Name", out object? value, out _);

            Assert.Multiple(() => {
                Assert.That(isFound, Is.True);
                Assert.That(value, Is.EqualTo("Indexed"));
            });
        }

        [Test]
        public void TryGetPathValue_PropertyHiddenByNew_ShouldUseMostDerivedProperty() {
            DerivedRecord derivedRecord = new() { Code = "Derived" };

            bool isFound = DictionaryUtils.TryGetPathValue(derivedRecord, "code", out object? value, out _);

            Assert.Multiple(() => {
                Assert.That(isFound, Is.True);
                Assert.That(value, Is.EqualTo("Derived"));
            });
        }

        [Test]
        public void TryGetPathValue_SegmentMatchesPropertiesDifferingByCase_ShouldPreferExactMatch() {
            CaseVariantRecord caseVariantRecord = new() { Value = "Upper", value = "Lower" };

            DictionaryUtils.TryGetPathValue(caseVariantRecord, "value", out object? value, out _);

            Assert.That(value, Is.EqualTo("Lower"));
        }

        [Test]
        public void TryGetPathValue_SegmentIsAmbiguousIgnoringCase_ShouldThrowArgumentException() {
            CaseVariantRecord caseVariantRecord = new() { Value = "Upper", value = "Lower" };

            Assert.That(
                () => DictionaryUtils.TryGetPathValue(caseVariantRecord, "VALUE", out _, out _),
                Throws.ArgumentException.With.Message.Contains("ambiguous")
            );
        }

        private class TestRecord {
            public int Id { get; set; }
            public string? Name { get; set; }
            public Address? Address { get; set; }
        }

        private class IndexedRecord {
            private readonly Dictionary<string, string> values = [];

            public string? Name { get; set; }

            public string this[string key] {
                get => values[key];
                set => values[key] = value;
            }
        }

        private class BaseRecord {
            public int Code { get; set; }
        }

        private class DerivedRecord : BaseRecord {
            public new string? Code { get; set; }
        }

        private class CaseVariantRecord {
            public string? Value { get; set; }

#pragma warning disable IDE1006 // 刻意建立僅大小寫不同的屬性以驗證比對規則
            public string? value { get; set; }
#pragma warning restore IDE1006
        }

        private class Address {
            public string? Street { get; set; }
            public string? City { get; set; }
            public string? Country { get; set; }
        }
    }
}
