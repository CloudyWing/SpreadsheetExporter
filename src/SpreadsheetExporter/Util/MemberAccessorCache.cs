using System;
using System.Linq.Expressions;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace CloudyWing.SpreadsheetExporter.Util;

/// <summary>
/// Provides cached getter delegates for instance properties and fields.
/// </summary>
internal static class MemberAccessorCache {
    // 以 ConditionalWeakTable 保存，避免 collectible assembly 卸載後仍被快取強參考
    private static readonly ConditionalWeakTable<MemberInfo, Func<object, object?>> Getters = new();

    /// <summary>
    /// Gets a cached getter delegate for the specified instance property or field.
    /// </summary>
    /// <param name="member">The property or field.</param>
    /// <returns>A delegate that reads the member value from a target instance.</returns>
    /// <exception cref="ArgumentException"><paramref name="member"/> is not an instance property or field.</exception>
    internal static Func<object, object?> GetGetter(MemberInfo member) {
        ArgumentNullException.ThrowIfNull(member);
        return Getters.GetValue(member, CreateGetter);
    }

    private static Func<object, object?> CreateGetter(MemberInfo member) {
        Type declaringType = member.DeclaringType
            ?? throw new ArgumentException($"Member '{member.Name}' has no declaring type.", nameof(member));

        ParameterExpression target = Expression.Parameter(typeof(object), "target");
        Expression instance = Expression.Convert(target, declaringType);
        Expression access = member switch {
            PropertyInfo property => Expression.Property(instance, property),
            FieldInfo field => Expression.Field(instance, field),
            _ => throw new ArgumentException($"Member '{member.Name}' must be a property or field.", nameof(member))
        };

        return Expression
            .Lambda<Func<object, object?>>(Expression.Convert(access, typeof(object)), target)
            .Compile();
    }
}
