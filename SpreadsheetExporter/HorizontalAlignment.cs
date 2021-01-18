using System.Text.Json.Serialization;

namespace CloudyWing.SpreadsheetExporter {
    [JsonConverter(typeof(JsonStringEnumConverter))]
    public enum HorizontalAlignment {
        None = 0,
        Left = 1,
        Center = 2,
        Right = 3,
        Justify = 4
    }
}
