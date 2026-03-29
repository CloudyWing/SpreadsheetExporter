namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a method that generates a value for a cell at the specified column and row position.
/// </summary>
/// <typeparam name="T">The type of value produced by this generator.</typeparam>
/// <param name="cellIndex">The zero-based column index of the cell.</param>
/// <param name="rowIndex">The zero-based row index of the cell.</param>
/// <returns>The generated value.</returns>
public delegate T CellGenerator<out T>(int cellIndex, int rowIndex);
