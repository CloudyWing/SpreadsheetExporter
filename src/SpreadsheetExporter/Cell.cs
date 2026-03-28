using System.Drawing;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// The spreadsheet cell.
/// </summary>
public class Cell {
    /// <summary>
    /// Gets or sets the cell content value generator. Pass the cell index and row index to the generator. The  index start at 0.
    /// </summary>
    /// <value>
    /// The cell content value.
    /// </value>
    public CellGenerator<object?>? ValueGenerator { get; set; }

    /// <summary>
    /// Gets or sets the cell style generator. Pass the cell index and row index to the generator. The  index start at 0.
    /// </summary>
    /// <value>
    /// The cell style.
    /// </value>
    public CellGenerator<CellStyle>? CellStyleGenerator { get; set; }

    /// <summary>
    /// Gets or sets the formula generator. Pass the cell index and row index to the generator. The  index start at 0.
    /// </summary>
    /// <value>
    /// The formula generator.
    /// </value>
    public CellGenerator<string?>? FormulaGenerator { get; set; }

    /// <summary>
    /// Gets or sets the data validation generator. Pass the cell index and row index to the generator. The  index start at 0.
    /// </summary>
    /// <value>
    /// The data validation generator.
    /// </value>
    public CellGenerator<DataValidation?>? DataValidationGenerator { get; set; }

    /// <summary>
    /// Gets or sets the zero-based column (X) and row (Y) position of this cell.
    /// </summary>
    /// <value>
    /// The zero-based column and row coordinates.
    /// </value>
    public Point Point { get; set; } // 本來考慮把 setter 改成 internal，但考量有開放自訂 Template，所以還是維持 public

    /// <summary>
    /// Gets or sets the cell size.
    /// </summary>
    /// <value>
    /// The cell size.
    /// </value>
    public Size Size { get; set; } = new Size(1, 1);

    /// <summary>
    /// Gets the content value.
    /// </summary>
    /// <returns>The content value.</returns>
    public object? GetValue() {
        return ValueGenerator?.Invoke(Point.X, Point.Y);
    }

    /// <summary>
    /// Gets the cell style.
    /// </summary>
    /// <returns>The cell style.</returns>
    public CellStyle GetCellStyle() {
        return CellStyleGenerator?.Invoke(Point.X, Point.Y) ?? CellStyle.Empty;
    }

    /// <summary>
    /// Gets the formula. If a formula starts with <c>=</c>, the prefix is automatically stripped.
    /// </summary>
    /// <returns>The formula string without the leading <c>=</c>, or <see langword="null"/> if no formula is set.</returns>
    public string? GetFormula() {
        return FormulaGenerator?
            .Invoke(Point.X, Point.Y)?
            .TrimStart(' ')
            .TrimStart('=');
    }

    /// <summary>
    /// Gets the data validation.
    /// </summary>
    /// <returns>The data validation, or <c>null</c> if no validation is specified.</returns>
    public DataValidation? GetDataValidation() {
        return DataValidationGenerator?.Invoke(Point.X, Point.Y);
    }

    /// <summary>
    /// Creates a shallow copy of this cell.
    /// </summary>
    /// <returns>A new <see cref="Cell"/> instance with the same field values.</returns>
    public Cell ShallowCopy() {
        return (Cell)MemberwiseClone();
    }
}
