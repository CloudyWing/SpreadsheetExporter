using System;
using System.Collections.Generic;
using System.Linq.Expressions;
using CloudyWing.SpreadsheetExporter.Util;

namespace CloudyWing.SpreadsheetExporter.Templates.RecordSet {
    /// <summary>The data column.</summary>
    /// <typeparam name="TField">The type of the record field.</typeparam>
    /// <typeparam name="TRecord">The type of the record.</typeparam>
    internal class DataColumn<TField, TRecord> : DataColumnBase<TRecord> {
        private readonly Dictionary<RecordContext<TRecord>, FieldContext<TField, TRecord>> contextMaps = new();

        /// <summary>Initializes a new instance of the <see cref="DataColumn{TField, TRecord}" /> class.</summary>
        /// <param name="fieldKey">The field key.</param>
        public DataColumn(string fieldKey) {
            FieldKey = fieldKey;
        }


        /// <summary>Initializes a new instance of the <see cref="DataColumn{TField, TRecord}" /> class.</summary>
        /// <param name="fieldKeyExpression">The field key expression.</param>
        public DataColumn(Expression<Func<TRecord, TField>> fieldKeyExpression) {
            FieldKey = GetFieldKeyByExpression(fieldKeyExpression);
        }

        /// <summary>Gets the field key.</summary>
        /// <value>The field key.</value>
        public string FieldKey { get; }

        /// <summary>Gets the field value generator.</summary>
        /// <value>The field value generator.</value>
        public Func<FieldContext<TField, TRecord>, object> FieldValueGenerator { get; set; }

        /// <summary>Gets or sets the field formula generator.</summary>
        /// <value>The field formula generator.</value>
        public Func<FieldContext<TField, TRecord>, string> FieldFormulaGenerator { get; set; }

        /// <summary>Gets or sets the field style generator.</summary>
        /// <value>The field style generator.</value>
        public Func<FieldContext<TField, TRecord>, CellStyle> FieldStyleGenerator { get; set; }

        private string GetFieldKeyByExpression(Expression<Func<TRecord, TField>> expression) {
            List<string> keys = new();
            if (expression is LambdaExpression lambda && lambda.Body is ConstantExpression constant) {
                keys.Add(constant.Value as string);
            } else {
                MemberExpression memberExpression = GetMemberExpression(expression);
                if (memberExpression == null) {
                    throw new ArgumentException("Wrong expression.", nameof(expression));
                }

                do {
                    keys.Add(memberExpression.Member.Name);
                    memberExpression = GetMemberExpression(memberExpression.Expression);
                } while (memberExpression != null);
            }

            keys.Reverse();
            return string.Join(".", keys);
        }

        private MemberExpression GetMemberExpression(Expression expression) {
            if (expression is null) {
                throw new ArgumentNullException(nameof(expression));
            }

            if (expression is MemberExpression member) {
                return member;
            } else if (expression is LambdaExpression lambda) {
                // 如果是 Value Type 的話 Body 會是 UnaryExpression
                // Reference Type 才會是直接取得到 MemberExpression
                return lambda.Body as MemberExpression
                    ?? (lambda.Body is UnaryExpression unary ? unary.Operand as MemberExpression : null);
            }

            return null;
        }

        public override object GetFieldValue(RecordContext<TRecord> recordContext) {
            if (string.IsNullOrWhiteSpace(FieldKey)) {
                return FieldValueGenerator is null ? "" : FieldValueGenerator(default);
            }

            FieldContext<TField, TRecord> fieldContext = GetFieldContextFromRecordContext(recordContext);

            return FieldValueGenerator is null ? fieldContext.Value : FieldValueGenerator(fieldContext);
        }

        private TResult GetFromGenerator<TResult>(
            Func<FieldContext<TField, TRecord>, TResult> generator, TResult defaultResult, RecordContext<TRecord> recordContext
        ) {
            if (generator is null) {
                return defaultResult;
            }

            if (string.IsNullOrWhiteSpace(FieldKey)) {
                return generator(default);
            }

            FieldContext<TField, TRecord> fieldContext = GetFieldContextFromRecordContext(recordContext);

            return generator(fieldContext);
        }

        private FieldContext<TField, TRecord> GetFieldContextFromRecordContext(RecordContext<TRecord> recordContext) {
            if (contextMaps.ContainsKey(recordContext)) {
                return contextMaps[recordContext];
            }

            int maxNestedLevel = FieldKey.Split('.').Length;

            IDictionary<string, object> maps = DictionaryUtils.ConvertFrom(recordContext.Record, maxNestedLevel);

            if (!maps.ContainsKey(FieldKey)) {
                throw new ArgumentException($"Data source does not contain property '{FieldKey}'.");
            }

            TField value = ChangeFieldValueType(maps[FieldKey]);


            return new FieldContext<TField, TRecord>(recordContext, FieldKey, value);
        }

        private TField ChangeFieldValueType(object value) {
            if (value is null) {
                return default;
            }

            Type fromType = Nullable.GetUnderlyingType(value.GetType()) ?? value.GetType();
            Type toType = Nullable.GetUnderlyingType(typeof(TField)) ?? typeof(TField);

            return toType.IsPrimitive && typeof(IConvertible).IsAssignableFrom(fromType)
                ? (TField)Convert.ChangeType(value, toType)
                : (TField)value;
        }

        public override string GetFieldFormula(RecordContext<TRecord> recordContext) {
            return GetFromGenerator(FieldFormulaGenerator, null, recordContext);
        }

        public override CellStyle GetFieldStyle(RecordContext<TRecord> recordContext) {
            return GetFromGenerator(FieldStyleGenerator, SpreadsheetManager.DefaultCellStyles.FieldStyle, recordContext);
        }
    }
}
