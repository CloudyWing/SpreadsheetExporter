namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Represents a zero-based rectangular range occupied by a spreadsheet cell.
/// </summary>
/// <param name="Column">The zero-based starting column.</param>
/// <param name="Row">The zero-based starting row.</param>
/// <param name="ColumnSpan">The number of columns occupied by the range.</param>
/// <param name="RowSpan">The number of rows occupied by the range.</param>
public readonly record struct CellRange(int Column, int Row, int ColumnSpan, int RowSpan) {
    /// <summary>
    /// Gets the zero-based ending column.
    /// </summary>
    public int LastColumn => Column + ColumnSpan - 1;

    /// <summary>
    /// Gets the zero-based ending row.
    /// </summary>
    public int LastRow => Row + RowSpan - 1;

    /// <summary>
    /// Gets a value indicating whether this range has valid coordinates and size.
    /// </summary>
    public bool IsValid => Column >= 0
        && Row >= 0
        && ColumnSpan > 0
        && RowSpan > 0;

    /// <summary>
    /// Determines whether this range intersects with another range.
    /// </summary>
    /// <param name="other">The other range.</param>
    /// <returns>
    /// <see langword="true" /> if the ranges intersect; otherwise, <see langword="false" />.
    /// </returns>
    public bool Intersects(CellRange other) {
        if (!IsValid || !other.IsValid) {
            return false;
        }

        return Column <= other.LastColumn
            && LastColumn >= other.Column
            && Row <= other.LastRow
            && LastRow >= other.Row;
    }
}
