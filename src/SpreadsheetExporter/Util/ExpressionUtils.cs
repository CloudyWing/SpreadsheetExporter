using System;
using System.Linq.Expressions;

namespace CloudyWing.SpreadsheetExporter.Util {
    internal static class ExpressionUtils {
        public static string GetMemberName<T>(Expression<Func<T, object>> expression) {
            if (expression is null) {
                throw new ArgumentNullException(nameof(expression));
            }

            string memberName = GetMemberName(expression.Body);

            if (memberName is null) {
                throw new ArgumentException("Expression格式錯誤。", nameof(expression));
            }

            return memberName;
        }

        private static string GetMemberName(Expression expression) {
            if (expression is null) {
                return null;
            }

            if (expression is ConstantExpression constantExpression) {
                return constantExpression.Value.ToString();
            }

            MemberExpression memberExpression = GetMemberExpression(expression);

            if (memberExpression is null) {
                return null;
            }

            string child = memberExpression.Member.Name;
            string parent = GetMemberName(memberExpression.Expression);

            return parent == null
                ? child
                : parent + "." + child;
        }

        private static MemberExpression GetMemberExpression(Expression expression) {
            // 如果是Value Type的話Body會是UnaryExpression
            // Reference Type才會是直接取得到MemberExpression
            return expression as MemberExpression
                ?? (expression is UnaryExpression unary ? unary.Operand as MemberExpression : null);
        }
    }
}
