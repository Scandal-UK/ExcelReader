// <copyright file="DictionaryMapper.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReader.Helpers;

using System.Linq.Expressions;
using System.Reflection;
using ExcelDataReader.Models;

/// <summary> Provides utility methods for mapping dictionaries to objects. </summary>
public static class DictionaryMapper
{
    private const string ItemCallMethod = "get_Item";

    private static readonly char[] Separator = [','];

    private static Type[] ParsableTypes { get; } =
    [
        typeof(int),
        typeof(long),
        typeof(short),
        typeof(double),
        typeof(byte),
        typeof(float),
        typeof(decimal),
        typeof(bool),
        typeof(DateTime),
        typeof(DateTimeOffset),
        typeof(Guid),
    ];

    /// <summary> Generates a mapping function that maps a dictionary to an object based on the provided column headers. </summary>
    /// <typeparam name="T">The target object type.</typeparam>
    /// <param name="columnHeaders">The list of column headers that define the mapping, for case-insensitive matching.</param>
    /// <returns>A mapping function that transforms dictionaries into objects of type <typeparamref name="T"/>.</returns>
    public static Func<IDictionary<string, string>, DictionaryMapperResult<T>> GenerateObjectMappingFunction<T>(List<string> columnHeaders)
    {
        var parameter = Expression.Parameter(typeof(IDictionary<string, string>), null);
        return Expression.Lambda<Func<IDictionary<string, string>, DictionaryMapperResult<T>>>(
            Expression.MemberInit(
                Expression.New(typeof(DictionaryMapperResult<T>)),
                Expression.Bind(
                    typeof(DictionaryMapperResult<T>).GetProperty(nameof(DictionaryMapperResult<T>.ObjectResult)),
                    Expression.MemberInit(Expression.New(typeof(T)), CreatePropertyAssignments<T>(columnHeaders, parameter))),
                Expression.Bind(
                    typeof(DictionaryMapperResult<T>).GetProperty(nameof(DictionaryMapperResult<T>.ExtraProperties)),
                    CreateUnmappedPropertiesDictionary(columnHeaders, parameter, typeof(T))),
                Expression.Bind(
                    typeof(DictionaryMapperResult<T>).GetProperty(nameof(DictionaryMapperResult<T>.Warnings)),
                    CreateWarningsList(columnHeaders, parameter, typeof(T)))),
            parameter).Compile();
    }

    /// <summary> Helper method to parse values from a CSV to a generic List. </summary>
    /// <typeparam name="T">Type to convert values to.</typeparam>
    /// <param name="value">CSV to parse.</param>
    /// <returns>Instance of <see cref="List{T}"/>.</returns>
    public static List<T> ParseList<T>(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return [];
        }

        var elements = value.Split(Separator, StringSplitOptions.RemoveEmptyEntries);
        return elements.Select(element => (T)Convert.ChangeType(element.Trim(), typeof(T))).ToList();
    }

    private static IEnumerable<MemberBinding> CreatePropertyAssignments<T>(List<string> columnHeaders, ParameterExpression parameter) =>
        typeof(T).GetProperties()
            .Select(propertyInfo => CreatePropertyAssignment<T>(columnHeaders, parameter, propertyInfo))
            .Where(propertyAssignment => propertyAssignment != null);

    private static MemberAssignment CreatePropertyAssignment<T>(List<string> columnHeaders, ParameterExpression parameter, PropertyInfo propertyInfo)
    {
        var columnName = columnHeaders.FirstOrDefault(c => string.Equals(c, propertyInfo.Name, StringComparison.OrdinalIgnoreCase));
        if (columnName == null)
        {
            return Expression.Bind(propertyInfo, Expression.Default(propertyInfo.PropertyType));
        }

        var keyExpression = Expression.Constant(columnName, typeof(string));
        var valueExpression = Expression.Condition(
           Expression.Call(parameter, nameof(IDictionary<string, string>.ContainsKey), Type.EmptyTypes, keyExpression),
           Expression.Call(parameter, ItemCallMethod, Type.EmptyTypes, keyExpression),
           Expression.Default(typeof(string)));

        return Expression.Bind(propertyInfo, CreateConversionExpression(valueExpression, propertyInfo.PropertyType));
    }

    private static BlockExpression CreateUnmappedPropertiesDictionary(List<string> columnHeaders, ParameterExpression parameter, Type targetType)
    {
        var dictionaryType = typeof(Dictionary<string, string>);
        var unmappedProperties = Expression.Variable(dictionaryType, nameof(DictionaryMapperResult<object>.ExtraProperties));

        return Expression.Block(
            new[] { unmappedProperties },
            Expression.Assign(unmappedProperties, Expression.New(dictionaryType)),
            Expression.Block(columnHeaders
               .Except(targetType.GetProperties().Select(prop => prop.Name), StringComparer.OrdinalIgnoreCase)
               .Select(column =>
               {
                   var keyExpression = Expression.Constant(column, typeof(string));
                   return Expression.IfThen(
                       Expression.Call(parameter, nameof(IDictionary<string, string>.ContainsKey), Type.EmptyTypes, keyExpression),
                       Expression.Call(
                           unmappedProperties,
                           dictionaryType.GetMethod(nameof(IDictionary<string, string>.Add), new[] { typeof(string), typeof(string) }),
                           keyExpression,
                           Expression.Call(parameter, ItemCallMethod, Type.EmptyTypes, keyExpression)));
               })),
            unmappedProperties);
    }

    private static BlockExpression CreateWarningsList(List<string> columnHeaders, ParameterExpression rowParameter, Type targetType)
    {
        var warnings = Expression.Variable(typeof(List<string>), nameof(DictionaryMapperResult<object>.Warnings));
        return Expression.Block(
            new[] { warnings },
            Expression.Assign(warnings, Expression.New(typeof(List<string>))),
            GenerateWarnings(columnHeaders, rowParameter, targetType, warnings),
            warnings);
    }

    private static BlockExpression GenerateWarnings(List<string> columnHeaders, ParameterExpression rowParameter, Type targetType, ParameterExpression warnings)
    {
        return Expression.Block(targetType.GetProperties()
            .Select(propertyInfo =>
            {
                var columnName = columnHeaders.FirstOrDefault(c => string.Equals(c, propertyInfo.Name, StringComparison.OrdinalIgnoreCase));
                if (columnName != null)
                {
                    var propertyType = Nullable.GetUnderlyingType(propertyInfo.PropertyType) ?? propertyInfo.PropertyType;
                    if (ParsableTypes.Contains(propertyType))
                    {
                        var valueExpression = Expression.Call(rowParameter, ItemCallMethod, Type.EmptyTypes, Expression.Constant(columnName, typeof(string)));
                        var tryParseCall = Expression.Call(
                            propertyType.GetMethod(nameof(int.TryParse), BindingFlags.Public | BindingFlags.Static, null, new[] { typeof(string), propertyType.MakeByRefType() }, null),
                            valueExpression,
                            Expression.Default(propertyType));

                        return (Expression)GenerateParseWarning(warnings, tryParseCall, propertyInfo, columnName, valueExpression);
                    }
                    else if (IsCustomClass(propertyType))
                    {
                        return Expression.Call(
                            warnings,
                            typeof(List<string>).GetMethod(nameof(List<string>.Add)),
                            Expression.Constant($"{columnName}: ignored - class property"));
                    }
                    else if (!IsListType(propertyType) && propertyType != typeof(string))
                    {
                        return Expression.Call(
                            warnings,
                            typeof(List<string>).GetMethod(nameof(List<string>.Add)),
                            Expression.Constant($"{columnName}: ignored - unsupported type"));
                    }
                }

                return null;
            })
            .Where(expr => expr != null)
            .ToList());
    }

    private static BlockExpression GenerateParseWarning(ParameterExpression warnings, MethodCallExpression tryParseCall, PropertyInfo propertyInfo, string columnName, Expression valueExpression)
    {
        var tryParseResult = Expression.Variable(typeof(bool), nameof(tryParseCall));
        return Expression.Block(
            new[] { tryParseResult },
            Expression.Assign(tryParseResult, tryParseCall),
            Expression.IfThen(
                Expression.IsFalse(tryParseResult),
                Expression.Call(
                    warnings,
                    typeof(List<string>).GetMethod(nameof(List<string>.Add)),
                    Expression.Constant($"{columnName}: cannot parse value '{valueExpression}' to type {propertyInfo.PropertyType}"))));
    }

    private static Expression CreateConversionExpression(Expression valueExpression, Type targetType)
    {
        if (ParsableTypes.Contains(Nullable.GetUnderlyingType(targetType) ?? targetType))
        {
            return GetTryParseExpressionByType(valueExpression, Nullable.GetUnderlyingType(targetType) ?? targetType);
        }
        else if (IsListType(targetType))
        {
            return HandleListType(valueExpression, targetType);
        }
        else if (IsCustomClass(targetType))
        {
            return Expression.Constant(null, targetType);
        }
        else if (targetType == typeof(string))
        {
            return Expression.Call(valueExpression, typeof(string).GetMethod(nameof(string.Trim), Type.EmptyTypes));
        }

        return Expression.New(targetType);
    }

    private static bool IsListType(Type targetType) => targetType.IsGenericType && targetType.GetGenericTypeDefinition() == typeof(List<>);

    private static bool IsCustomClass(Type targetType) => targetType.IsClass && !targetType.Namespace.StartsWith(nameof(System));

    private static ConditionalExpression HandleListType(Expression valueExpression, Type targetType) =>
        Expression.Condition(
            Expression.Equal(valueExpression, Expression.Constant(string.Empty, typeof(string))),
            Expression.New(typeof(List<>).MakeGenericType(targetType.GenericTypeArguments[0])),
            Expression.Call(
                typeof(DictionaryMapper).GetMethod(nameof(ParseList)).MakeGenericMethod(targetType.GenericTypeArguments[0]),
                valueExpression));

    private static BlockExpression GetTryParseExpressionByType(Expression valueExpression, Type targetType)
    {
        var parsedVariable = Expression.Variable(targetType, nameof(targetType));
        return Expression.Block(
            new[] { parsedVariable },
            Expression.IfThen(
                Expression.Call(
                    targetType,
                    nameof(int.TryParse),
                    null,
                    new Expression[] { valueExpression, parsedVariable }),
                Expression.Empty()),
            parsedVariable);
    }
}
