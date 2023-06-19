#### [CloudyWing.SpreadsheetExporter](index.md 'index')
### [CloudyWing.SpreadsheetExporter](CloudyWing.SpreadsheetExporter.md 'CloudyWing.SpreadsheetExporter')

## SpreadsheetFileMode Enum

The file mode of spreadsheet

```csharp
public enum SpreadsheetFileMode
```
### Fields

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.Create'></a>

`Create` 0

Create a file. If the file already exists, it will be overwritten.

<a name='CloudyWing.SpreadsheetExporter.SpreadsheetFileMode.CreateNew'></a>

`CreateNew` 1

Create a new file. If the file already exists, an IOException exception is thrown.