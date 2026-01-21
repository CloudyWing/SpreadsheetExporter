#### [CloudyWing\.SpreadsheetExporter](index.md 'index')
### [CloudyWing\.SpreadsheetExporter\.Extensions](CloudyWing.SpreadsheetExporter.Extensions.md 'CloudyWing\.SpreadsheetExporter\.Extensions')

## EnumerableExtensions Class

Provides extension methods for [System\.Collections\.Generic\.IEnumerable&lt;&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')\.

```csharp
internal static class EnumerableExtensions
```

Inheritance [System\.Object](https://learn.microsoft.com/en-us/dotnet/api/system.object 'System\.Object') &#129106; EnumerableExtensions
### Methods

<a name='CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.AsReadOnly_T_(thisSystem.Collections.Generic.IEnumerable_T_)'></a>

## EnumerableExtensions\.AsReadOnly\<T\>\(this IEnumerable\<T\>\) Method

Converts the specified [System\.Collections\.Generic\.IEnumerable&lt;&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1') to a read\-only list\.

```csharp
public static System.Collections.Generic.IReadOnlyList<T> AsReadOnly<T>(this System.Collections.Generic.IEnumerable<T> enumerable);
```
#### Type parameters

<a name='CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.AsReadOnly_T_(thisSystem.Collections.Generic.IEnumerable_T_).T'></a>

`T`

The type of the elements in the collection\.
#### Parameters

<a name='CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.AsReadOnly_T_(thisSystem.Collections.Generic.IEnumerable_T_).enumerable'></a>

`enumerable` [System\.Collections\.Generic\.IEnumerable&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')[T](CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.md#CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.AsReadOnly_T_(thisSystem.Collections.Generic.IEnumerable_T_).T 'CloudyWing\.SpreadsheetExporter\.Extensions\.EnumerableExtensions\.AsReadOnly\<T\>\(this System\.Collections\.Generic\.IEnumerable\<T\>\)\.T')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ienumerable-1 'System\.Collections\.Generic\.IEnumerable\`1')

The enumerable to convert\.

#### Returns
[System\.Collections\.Generic\.IReadOnlyList&lt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')[T](CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.md#CloudyWing.SpreadsheetExporter.Extensions.EnumerableExtensions.AsReadOnly_T_(thisSystem.Collections.Generic.IEnumerable_T_).T 'CloudyWing\.SpreadsheetExporter\.Extensions\.EnumerableExtensions\.AsReadOnly\<T\>\(this System\.Collections\.Generic\.IEnumerable\<T\>\)\.T')[&gt;](https://learn.microsoft.com/en-us/dotnet/api/system.collections.generic.ireadonlylist-1 'System\.Collections\.Generic\.IReadOnlyList\`1')  
A read\-only list containing the elements of the enumerable\.