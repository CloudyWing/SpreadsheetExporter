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
            IDictionary<string, object> result = DictionaryUtils.ConvertFrom(record, 1);

            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result.ContainsKey("Id"), Is.True);
            Assert.That(result["Id"], Is.EqualTo(1));
            Assert.That(result.ContainsKey("Name"), Is.True);
            Assert.That(result["Name"], Is.EqualTo("Test"));
            Assert.That(result.ContainsKey("Address"), Is.True);
            Assert.That(result["Address"], Is.EqualTo(record.Address));
        }

        [Test]
        public void ConvertFrom_NestedPropertyLevelIs2_ShouldIncludeNestedProperties() {
            IDictionary<string, object> result = DictionaryUtils.ConvertFrom(record, 2);

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
        }

        private class TestRecord {
            public int Id { get; set; }
            public string? Name { get; set; }
            public Address? Address { get; set; }
        }

        private class Address {
            public string? Street { get; set; }
            public string? City { get; set; }
            public string? Country { get; set; }
        }
    }
}
