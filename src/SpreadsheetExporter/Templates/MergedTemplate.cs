using System;
using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter.Templates {
    /// <summary>
    /// The merged template. Merge templates into new template.
    /// </summary>
    /// <seealso cref="ITemplate" />
    public class MergedTemplate : ITemplate {
        private readonly TemplateContext templateContext;

        /// <summary>Initializes a new instance of the <see cref="MergedTemplate" /> class.</summary>
        /// <param name="templates">The templates.</param>
        /// <exception cref="ArgumentNullException">templates</exception>
        public MergedTemplate(IEnumerable<ITemplate> templates) {
            if (templates == null) {
                throw new ArgumentNullException(nameof(templates));
            }
            templateContext = TemplateContext.Create(templates);
        }

        /// <summary>Initializes a new instance of the <see cref="MergedTemplate" /> class.</summary>
        /// <param name="templates">The templates.</param>
        public MergedTemplate(params ITemplate[] templates) : this(templates as IEnumerable<ITemplate>) { }

        /// <summary>Gets the height of rows.</summary>
        /// <value>The height of rows.</value>
        public IReadOnlyDictionary<int, double> RowHeights => templateContext.RowHeights;

        /// <summary>Gets the context.</summary>
        /// <returns>The template context.</returns>
        public TemplateContext GetContext() {
            return templateContext;
        }
    }
}
