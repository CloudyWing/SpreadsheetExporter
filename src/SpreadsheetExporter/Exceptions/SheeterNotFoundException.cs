using System;
using System.Runtime.Serialization;

namespace CloudyWing.SpreadsheetExporter.Exceptions {
    public class SheeterNotFoundException : Exception {
        public SheeterNotFoundException() : this("Could not find Sheeter.") { }

        public SheeterNotFoundException(string message) : base(message) {
        }

        public SheeterNotFoundException(string message, Exception innerException) : base(message, innerException) {
        }

        protected SheeterNotFoundException(SerializationInfo info, StreamingContext context) : base(info, context) {
        }
    }
}
