using System;

namespace CloudyWing.SpreadsheetExporter.Templates.Json;

internal static class JsonParseExceptionFactory {
    internal static FormatException InvalidType(JsonParseContext context, string expectedType) {
        return new FormatException($"{JsonDiagnosticCodes.InvalidType}: {context.Path} must be {expectedType}.");
    }

    internal static FormatException MissingRequiredProperty(JsonParseContext context) {
        return new FormatException($"{JsonDiagnosticCodes.MissingRequiredProperty}: {context.Path} is required.");
    }

    internal static FormatException InvalidValue(JsonParseContext context, string message) {
        return new FormatException($"{JsonDiagnosticCodes.InvalidValue}: {context.Path} {message}");
    }

    internal static InvalidOperationException InvalidOperation(JsonParseContext context, string message) {
        return new InvalidOperationException($"{JsonDiagnosticCodes.InvalidValue}: {context.Path} {message}");
    }

    internal static NotSupportedException NotSupported(JsonParseContext context, string message) {
        return new NotSupportedException($"{JsonDiagnosticCodes.InvalidValue}: {context.Path} {message}");
    }
}
