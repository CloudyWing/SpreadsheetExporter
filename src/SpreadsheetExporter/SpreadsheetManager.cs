using System;
using System.Diagnostics.CodeAnalysis;
using System.Threading;
using CloudyWing.SpreadsheetExporter.Config;

namespace CloudyWing.SpreadsheetExporter;

/// <summary>
/// Provides global configuration and factory methods for creating spreadsheet documents.
/// </summary>
public static class SpreadsheetManager {
#if NET10_0_OR_GREATER
    private static readonly Lock RendererFactoryLock = new();
    private static readonly Lock CellStyleLock = new();
#else
    private static readonly object RendererFactoryLock = new();
    private static readonly object CellStyleLock = new();
#endif
    private static Func<ISpreadsheetRenderer>? rendererFactory;

    /// <summary>
    /// Gets or sets the default cell styles configuration.
    /// Setting this overrides the built-in defaults.
    /// </summary>
    [AllowNull]
    public static CellStyleConfiguration DefaultCellStyles {
        get {
            if (field is null) {
                lock (CellStyleLock) {
                    field ??= new CellStyleConfiguration();
                }
            }
            return field;
        }
        set {
            lock (CellStyleLock) {
                field = value;
            }
        }
    }

    /// <summary>
    /// Sets the renderer factory used to create renderer instances.
    /// </summary>
    /// <param name="rendererFactory">The factory function that returns a new <see cref="ISpreadsheetRenderer"/>.</param>
    /// <exception cref="ArgumentNullException"><paramref name="rendererFactory"/> is <see langword="null"/>.</exception>
    /// <exception cref="ArgumentException">The factory returns <see langword="null"/>.</exception>
    public static void SetRenderer(Func<ISpreadsheetRenderer> rendererFactory) {
        ArgumentNullException.ThrowIfNull(rendererFactory);

        lock (RendererFactoryLock) {
            if (rendererFactory() is null) {
                throw new ArgumentException(
                    "Factory return value cannot be null.", nameof(rendererFactory)
                );
            }

            SpreadsheetManager.rendererFactory = rendererFactory;
        }
    }

    /// <summary>
    /// Creates a new <see cref="SpreadsheetDocument"/> using the registered renderer factory.
    /// </summary>
    /// <returns>A new <see cref="SpreadsheetDocument"/> instance.</returns>
    /// <exception cref="InvalidOperationException">The renderer factory has not been set via <see cref="SetRenderer"/>.</exception>
    public static SpreadsheetDocument CreateDocument() {
        return rendererFactory is null
            ? throw new InvalidOperationException("Renderer factory is not set. Call SetRenderer first.")
            : new SpreadsheetDocument(rendererFactory());
    }
}
