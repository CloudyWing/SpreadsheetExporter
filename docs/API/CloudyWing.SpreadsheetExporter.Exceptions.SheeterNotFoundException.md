#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Exceptions](CloudyWing.SpreadsheetExporter.Exceptions.md 'CloudyWing\.SpreadsheetExporter\.Exceptions')

## SheeterNotFoundException Class

The exception for sheeter not found\.

```csharp
public class SheeterNotFoundException : System.Exception
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; [System\.Exception](https://learn.microsoft.com/en-us/dotnet/api/system.exception 'System\.Exception') &#129106; SheeterNotFoundException

### See Also
- [System\.Exception](https://learn.microsoft.com/en-us/dotnet/api/system.exception 'System\.Exception')
### Constructors

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException()'></a>

## SheeterNotFoundException\(\) Constructor

Initializes a new instance of the [SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException') class\.

```csharp
public SheeterNotFoundException();
```

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(string)'></a>

## SheeterNotFoundException\(string\) Constructor

Initializes a new instance of the [SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException') class\.

```csharp
public SheeterNotFoundException(string message);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(string).message'></a>

`message` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The message that describes the error\.

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(string,System.Exception)'></a>

## SheeterNotFoundException\(string, Exception\) Constructor

Initializes a new instance of the [SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException') class\.

```csharp
public SheeterNotFoundException(string message, System.Exception innerException);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(string,System.Exception).message'></a>

`message` [System\.String](https://learn.microsoft.com/en-us/dotnet/api/system.string 'System\.String')

The error message that explains the reason for the exception\.

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(string,System.Exception).innerException'></a>

`innerException` [System\.Exception](https://learn.microsoft.com/en-us/dotnet/api/system.exception 'System\.Exception')

The exception that is thecause of the current exception, or a null reference \(\<span class="keyword"\>
  \<span class="languageSpecificText"\>
    \<span class="cs"\>null\</span\>
    \<span class="vb"\>Nothing\</span\>
    \<span class="cpp"\>nullptr\</span\>
  \</span\>
\</span\>\<span class="nu"\>a null reference \(\<span class="keyword"\>Nothing\</span\> in Visual Basic\)\</span\> in Visual Basic\) if no inner exception is specified\.

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(System.Runtime.Serialization.SerializationInfo,System.Runtime.Serialization.StreamingContext)'></a>

## SheeterNotFoundException\(SerializationInfo, StreamingContext\) Constructor

Initializes a new instance of the [SheeterNotFoundException](CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.md 'CloudyWing\.SpreadsheetExporter\.Exceptions\.SheeterNotFoundException') class\.

```csharp
protected SheeterNotFoundException(System.Runtime.Serialization.SerializationInfo info, System.Runtime.Serialization.StreamingContext context);
```
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(System.Runtime.Serialization.SerializationInfo,System.Runtime.Serialization.StreamingContext).info'></a>

`info` [System\.Runtime\.Serialization\.SerializationInfo](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.serializationinfo 'System\.Runtime\.Serialization\.SerializationInfo')

The [System\.Runtime\.Serialization\.SerializationInfo](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.serializationinfo 'System\.Runtime\.Serialization\.SerializationInfo') that holds the serialized object data about the exception being thrown\.

<a name='CloudyWing.SpreadsheetExporter.Exceptions.SheeterNotFoundException.SheeterNotFoundException(System.Runtime.Serialization.SerializationInfo,System.Runtime.Serialization.StreamingContext).context'></a>

`context` [System\.Runtime\.Serialization\.StreamingContext](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.streamingcontext 'System\.Runtime\.Serialization\.StreamingContext')

The [System\.Runtime\.Serialization\.StreamingContext](https://learn.microsoft.com/en-us/dotnet/api/system.runtime.serialization.streamingcontext 'System\.Runtime\.Serialization\.StreamingContext') that contains contextual information about the source or destination\.