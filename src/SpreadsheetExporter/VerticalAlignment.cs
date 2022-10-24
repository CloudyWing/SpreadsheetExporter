using System.Text.Json.Serialization;

namespace CloudyWing.SpreadsheetExporter {
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public enum VerticalAlignment {
        Top = 0,
        Middle = 1,
        Bottom = 2
    }
}
