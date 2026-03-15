using System.Collections.Concurrent;
using System.Collections.Frozen;
using System.Diagnostics.CodeAnalysis;
using System.Linq.Expressions;
using System.Reflection;

namespace StreamScheme.Mappers;

public interface IRowMapper
{
    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    IEnumerable<IEnumerable<FieldValue>> ToRows<T>(IEnumerable<T> items);
}

public class ReflectionRowMapper : IRowMapper
{
    [RequiresUnreferencedCode("Uses reflection and compiled expressions to map properties.")]
    [RequiresDynamicCode("Uses Expression.Compile() which is not compatible with AOT.")]
    public IEnumerable<IEnumerable<FieldValue>> ToRows<T>(IEnumerable<T> items)
    {
        ConcurrentDictionary<Type, FrozenDictionary<object, FieldValue>> enumCache = [];
        FieldValue falseFieldValue = false;
        FieldValue trueFieldValue = true;

        var accessors = PropertyAccessorCache<T>.Accessors;

        yield return accessors.Select(a => (FieldValue)a.Name).ToArray();

        foreach (var item in items)
        {
            var row = new FieldValue[accessors.Length];
            for (var i = 0; i < accessors.Length; i++)
            {
                row[i] = ToFieldValue(
                    accessors[i].GetValue(item),
                    falseFieldValue,
                    trueFieldValue,
                    enumCache);
            }

            yield return row;
        }
    }

    private static FieldValue ToFieldValue(
        object? value,
        FieldValue falseFieldValue,
        FieldValue trueFieldValue,
        ConcurrentDictionary<Type, FrozenDictionary<object, FieldValue>> enumCache) => value switch
        {
            null => FieldValue.EmptyField,
            string s => s,
            double d => d,
            float f => (double)f,
            int i => (double)i,
            long l => (double)l,
            short s => (double)s,
            byte b => (double)b,
            decimal m => (double)m,
            DateOnly date => date,
            bool b => b ? trueFieldValue : falseFieldValue,
            Enum e => ResolveEnum(e, enumCache),
            _ => value.ToString() ?? string.Empty
        };

    private static FieldValue ResolveEnum(Enum value, ConcurrentDictionary<Type, FrozenDictionary<object, FieldValue>> enumCache)
    {
        var cache = enumCache.GetOrAdd(
            value.GetType(),
            static type => Enum.GetValues(type)
                .Cast<object>()
                .ToDictionary(v => v, v => (FieldValue)v.ToString()!)
                .ToFrozenDictionary());

        return cache[value];
    }

    private static class PropertyAccessorCache<T>
    {
        public static readonly (string Name, Func<T, object?> GetValue)[] Accessors = typeof(T)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Where(property => property.CanRead)
            .Select(property => (property.Name, CompileGetter(property)))
            .ToArray();

        private static Func<T, object?> CompileGetter(PropertyInfo property)
        {
            var parameter = Expression.Parameter(typeof(T), "instance");
            var propertyAccess = Expression.Property(parameter, property);
            var boxed = Expression.Convert(propertyAccess, typeof(object));
            return Expression.Lambda<Func<T, object?>>(boxed, parameter).Compile();
        }
    }
}
