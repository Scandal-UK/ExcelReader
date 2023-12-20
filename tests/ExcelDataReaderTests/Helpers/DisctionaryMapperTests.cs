// <copyright file="DisctionaryMapperTests.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReaderTests.Helpers;

using ExcelDataReader.Helpers;
using ExcelDataReaderTests.Helpers.Stubs;

/// <summary> Tests for the <see cref="DictionaryMapper"/> class. </summary>
public class DictionaryMapperTests
{
    /// <summary> Test that mapping works with expected columns. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_ValidData_ShouldMapProperties()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Name), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Name", "John" },
            { "Age", "30" },
        };

        // Act
        var result = mappingFunction(row).ObjectResult;

        // Assert
        result.Should().NotBeNull();
        result.Id.Should().Be("1");
        result.Name.Should().Be("John");
        result.Age.Should().Be(30);
    }

    /// <summary> Test that mapping still works with unordered columns. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_UnorderedData_ShouldMapPropertiesCorrectly()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Age), nameof(FakeEntity.Name), nameof(FakeEntity.Id) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Name", "Jane" },
            { "Age", "25" },
            { "Id", "2" },
        };

        // Act
        var result = mappingFunction(row).ObjectResult;

        // Assert
        result.Should().NotBeNull();
        result.Id.Should().Be("2");
        result.Name.Should().Be("Jane");
        result.Age.Should().Be(25);
    }

    /// <summary> Test that mapping still works with missing columns. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MissingData_ShouldMapPropertiesWithDefaults()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "3" },
            { "Age", string.Empty },
        };

        // Act
        var result = mappingFunction(row).ObjectResult;

        // Assert
        result.Should().NotBeNull();
        result.Id.Should().Be("3");
        result.Age.Should().Be(0);
    }

    /// <summary> Test that mapping still works with incorrect casing on the header names. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_IncorrectCaseHeaders_ShouldMapToProperties()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id).ToLower(), nameof(FakeEntity.Name).ToLower(), nameof(FakeEntity.Age).ToLower() };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "id", "1" },
            { "name", "John" },
            { "age", "30" },
        };

        // Act
        var result = mappingFunction(row).ObjectResult;

        // Assert
        result.Should().NotBeNull();
        result.Id.Should().Be("1");
        result.Name.Should().Be("John");
        result.Age.Should().Be(30);
    }

    /// <summary> Test that mapping still works and collects unmapped properties. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_UnmappedProperties_ShouldGatherUnmappedProperties()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Name), nameof(FakeEntity.Age), "Foo" };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Name", "John" },
            { "Age", "30" },
            { "Foo", "Bar" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;
        var unmappedProperties = result.ExtraProperties;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Id.Should().Be("1");
        mappedItem.Name.Should().Be("John");
        mappedItem.Age.Should().Be(30);

        unmappedProperties.Should().ContainKey("Foo");
        unmappedProperties.Should().NotContainKey("Id");
        unmappedProperties.Should().NotContainKey("Name");
        unmappedProperties.Should().NotContainKey("Age");
    }

    /// <summary> Test that mapping fails with mismatched Integer type. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MismatchedIntType_ShouldGenerate_Warning()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.IsMember), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "IsMember", "true" },
            { "Age", "string value" },
        };

        // Act
        var result = mappingFunction(row);

        // Assert
        result.Warnings.Count.Should().Be(1);
    }

    /// <summary> Test that mapping fails with mismatched Boolean type. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MismatchedBoolType_ShouldGenerate_Warning()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.IsMember), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "IsMember", "string value" },
            { "Age", "30" },
        };

        // Act
        var result = mappingFunction(row);

        // Assert
        result.Warnings.Count.Should().Be(1);
    }

    /// <summary> Test that mapping fails with mismatched DateTimeOffset type. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MismatchedDateTimeOffsetType_ShouldGenerate_Warning()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.CreatedDate), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "CreatedDate", "string value" },
            { "Age", "30" },
        };

        // Act
        var result = mappingFunction(row);

        // Assert
        result.Warnings.Count.Should().Be(1);
    }

    /// <summary> Test that mapping fails with multiple mismatched types. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MismatchedColumns_ShouldGenerate_MultipleWarnings()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.IsMember), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "IsMember", "string value" },
            { "Age", "string value" },
        };

        // Act
        var result = mappingFunction(row);

        // Assert
        result.Warnings.Count.Should().Be(2);
    }

    /// <summary> Test that List{string} values are parsed. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_StringList_IsParsed()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Tags) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Tags", "first, second, third" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Tags.Count.Should().Be(3);
        mappedItem.Tags.First().Should().Be("first");
        mappedItem.Tags.Last().Should().Be("third");
    }

    /// <summary> Test that List{int} values are parsed. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_IntList_IsParsed()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Numbers) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Numbers", "1, 2, 3, 4" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Numbers.Count.Should().Be(4);
        mappedItem.Numbers.First().Should().Be(1);
        mappedItem.Numbers.Last().Should().Be(4);
    }

    /// <summary> Test that empty List{} values are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_EmptyListValues_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Tags), nameof(FakeEntity.Numbers), nameof(FakeEntity.RelatedFakeEntities) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Tags", string.Empty },
            { "Numbers", string.Empty },
            { "RelatedFakeEntities", string.Empty },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Tags.Count.Should().Be(0);
        mappedItem.Numbers.Count.Should().Be(0);
        mappedItem.RelatedFakeEntities.Count.Should().Be(0);
    }

    /// <summary> Test that for complex types, values are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_ComplexTypes_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.RelatedEntity) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "RelatedEntity", "Foo Bar" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.RelatedEntity.Should().BeNull();
        result.Warnings.Count.Should().Be(1); // Expect a warning to be logged
    }

    /// <summary> Test that missing properties are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_MissingProperties_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), "Foo", "Bar" };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Bar", "Something else" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Id.Should().Be("1");
        result.ExtraProperties.Count.Should().Be(1);
        result.ExtraProperties.Should().ContainKey("Bar");
    }

    /// <summary> Test that unmappable properties are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_UnmappableGenericTypes_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Unmappable) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Unmappable", "value" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Id.Should().Be("1");
        mappedItem.Unmappable.Should().BeEmpty();
        result.Warnings.Count.Should().Be(1);
    }

    /// <summary> Test that unparsable properties are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_UnparsableTypes_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Unparsable) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Unparsable", "value" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Id.Should().Be("1");
        mappedItem.Unparsable.Should().Be(default);
        result.Warnings.Count.Should().Be(1);
    }

    /// <summary> Test that nullable unparsable properties are ignored. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_NullableUnparsableTypes_AreIgnored()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.NullableUnparsable) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "NullableUnparsable", "value" },
        };

        // Act
        var result = mappingFunction(row);
        var mappedItem = result.ObjectResult;

        // Assert
        mappedItem.Should().NotBeNull();
        mappedItem.Id.Should().Be("1");
        mappedItem.NullableUnparsable.Should().BeNull();
    }

    /// <summary> Test that mapping works with additional whitespace. </summary>
    [Fact]
    public void GenerateObjectMappingFunction_WhiteSpace_ShouldBeTrimmed()
    {
        // Arrange
        var columnHeaders = new List<string> { nameof(FakeEntity.Id), nameof(FakeEntity.Name), nameof(FakeEntity.Age) };
        var mappingFunction = DictionaryMapper.GenerateObjectMappingFunction<FakeEntity>(columnHeaders);

        var row = new Dictionary<string, string>
        {
            { "Id", "1" },
            { "Name", "  John  " },
            { "Age", "  30  " },
        };

        // Act
        var result = mappingFunction(row).ObjectResult;

        // Assert
        result.Should().NotBeNull();
        result.Id.Should().Be("1");
        result.Name.Should().Be("John");
        result.Age.Should().Be(30);
    }
}
