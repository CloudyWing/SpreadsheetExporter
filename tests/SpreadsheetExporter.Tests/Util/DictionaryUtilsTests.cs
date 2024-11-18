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

            result.Count.Should().Be(3);
            result.Should().ContainKey("Id").WhoseValue.Should().Be(1);
            result.Should().ContainKey("Name").WhoseValue.Should().Be("Test");
            result.Should().ContainKey("Address").WhoseValue.Should().Be(record.Address);
        }

        [Test]
        public void ConvertFrom_NestedPropertyLevelIs2_ShouldIncludeNestedProperties() {
            IDictionary<string, object> result = DictionaryUtils.ConvertFrom(record, 2);

            result.Count.Should().Be(5);
            result.Should().ContainKey("Id").WhoseValue.Should().Be(1);
            result.Should().ContainKey("Name").WhoseValue.Should().Be("Test");
            result.Should().ContainKey("Address.Street").WhoseValue.Should().Be("Test Street");
            result.Should().ContainKey("Address.City").WhoseValue.Should().Be("Test City");
            result.Should().ContainKey("Address.Country").WhoseValue.Should().Be("Test Country");
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
