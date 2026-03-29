using System.Drawing;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet;

/// <summary>
/// The base class for record set column definitions.
/// </summary>
/// <typeparam name="T">The type of the record.</typeparam>
public abstract class RecordSetColumnBase<T> {
    /// <summary>
    /// Gets or sets the header text.
    /// </summary>
    /// <value>The header text.</value>
    public string? HeaderText { get; set; }

    /// <summary>
    /// Gets or sets the header style.
    /// </summary>
    /// <value>The header style.</value>
    public CellStyle HeaderStyle { get; set; }

    /// <summary>
    /// Gets the point.
    /// </summary>
    /// <value>The point.</value>
    public Point Point { get; internal set; }

    /// <summary>
    /// Gets the column span.
    /// </summary>
    /// <value>The column span.</value>
    public int ColumnSpan => ChildColumns.Count == 0 ? 1 : ChildColumns.ColumnSpan;

    /// <summary>
    /// Gets the row span.
    /// </summary>
    /// <value>The row span.</value>
    public int RowSpan => ParentCollection!.RowSpan - ChildColumns.RowSpan;

    /// <summary>
    /// Gets the child columns.
    /// </summary>
    /// <value>The child columns.</value>
    public RecordSetColumnCollection<T> ChildColumns {
        get => field ??= new RecordSetColumnCollection<T>(this);
        private set;
    }

    /// <summary>
    /// Gets the parent collection this column belongs to.
    /// </summary>
    /// <value>The parent collection.</value>
    public RecordSetColumnCollection<T>? ParentCollection { get; internal set; }

    /// <summary>
    /// Gets the header depth, representing the number of row layers this column and its children occupy.
    /// </summary>
    /// <value>The header depth.</value>
    public int HeaderDepth => ChildColumns.Count == 0
        ? 1 : ChildColumns.RowSpan + 1;

    /// <summary>
    /// Gets the field value for the specified context.
    /// </summary>
    /// <param name="context">The record context.</param>
    /// <returns>The field value, or <see langword="null"/> if no value is produced.</returns>
    public abstract object? GetFieldValue(RecordContext<T> context);

    /// <summary>
    /// Gets the field formula for the specified context.
    /// </summary>
    /// <param name="context">The record context.</param>
    /// <returns>The field formula, or <see langword="null"/> if no formula is specified.</returns>
    public abstract string? GetFieldFormula(RecordContext<T> context);

    /// <summary>
    /// Gets the field style for the specified context.
    /// </summary>
    /// <param name="context">The record context.</param>
    /// <returns>The field style.</returns>
    public abstract CellStyle GetFieldStyle(RecordContext<T> context);

    /// <summary>
    /// Gets the field data validation for the specified context.
    /// </summary>
    /// <param name="context">The record context.</param>
    /// <returns>The field data validation, or <see langword="null"/> if no validation is specified.</returns>
    public abstract DataValidation? GetFieldDataValidation(RecordContext<T> context);
}
