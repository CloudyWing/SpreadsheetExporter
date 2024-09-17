using System;
using System.Collections.Generic;

namespace CloudyWing.SpreadsheetExporter.Templates {
    /// <summary>
    /// The merged template. Merge templates into new template.
    /// </summary>
    /// <seealso cref="ITemplate" />
    /// <param name="templates">The templates.</param>
    /// <exception cref="ArgumentNullException">templates</exception>
    public class MergedTemplate(IEnumerable<ITemplate> templates) : ITemplate {
        private readonly IEnumerable<ITemplate> templates =
            templates ?? throw new ArgumentNullException(nameof(templates));

        /// <summary>
        /// Initializes a new instance of the <see cref="MergedTemplate" /> class.
        /// </summary>
        /// <param name="templates">The templates.</param>
        public MergedTemplate(params ITemplate[] templates)
            : this(templates as IEnumerable<ITemplate>) { }

        /// <summary>
        /// Gets the context.
        /// </summary>
        /// <returns>The template context.</returns>
        public TemplateContext GetContext() {
            return TemplateContext.Create(templates);
        }
    }
}
