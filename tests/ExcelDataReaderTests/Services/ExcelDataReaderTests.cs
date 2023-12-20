// <copyright file="ExcelDataReaderTests.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReaderTests.Services;

using System.Runtime.CompilerServices;
using ExcelDataReader.Services;
using global::ExcelDataReaderTests.Helpers.Stubs;

/// <summary> Tests to ensure that the ExcelDataReader class works as expected. </summary>
public class ExcelDataReaderTests
{
    /// <summary> Test that batch-reading maps columns to properties as expected. </summary>
    [Fact]
    public void BatchReadExcel_ValidData_ShouldMapColumnsToObjectProperties()
    {
        // Arrange
        var excelDataReader = new ExcelDataReader<FakeEntity>();
        var excelStream = GetTestSpreadsheet();

        // Act
        var result = excelDataReader.BatchReadExcel(excelStream, 10).ToList();

        // Assert
        result.Should().HaveCount(1);
        var resultList = result[0].ToList();
        resultList.Should().HaveCount(6);

        resultList[0].ObjectResult.Age.Should().Be(30);
        resultList[0].ObjectResult.Name.Should().Be("John");
        resultList[0].ObjectResult.IsMember.Should().BeTrue();

        resultList[0].ExtraProperties.Count.Should().Be(1);
        resultList[0].ExtraProperties.First().Key.Should().Be("Attribute-Foo");
        resultList[0].ExtraProperties.First().Value.Should().Be("Bar");

        resultList[2].ObjectResult.Age.Should().Be(40);
        resultList[2].ObjectResult.Name.Should().Be("Alice");
        resultList[2].ObjectResult.IsMember.Should().BeTrue();

        resultList[2].ExtraProperties.Count.Should().Be(1);
        resultList[2].ExtraProperties.First().Key.Should().Be("Attribute-Foo");
        resultList[2].ExtraProperties.First().Value.Should().Be("Bar");
    }

    /// <summary> Test that batch-reading continues to work across multiple iterations. </summary>
    [Fact]
    public void BatchReadExcel_MultipleIterations_ShouldReadAllRecords()
    {
        // Arrange
        var excelDataReader = new ExcelDataReader<FakeEntity>();
        var excelStream = GetTestSpreadsheet();

        // Act
        var result = excelDataReader.BatchReadExcel(excelStream, batchSize: 5).ToList();

        // Assert
        result.Should().HaveCount(2); // batches
        var resultList = result[0].ToList();
        resultList.Should().HaveCount(5);
        result[1].Should().HaveCount(1);

        resultList[0].ObjectResult.Id.Should().Be("1");
        resultList[0].ObjectResult.Age.Should().Be(30);
        resultList[0].ObjectResult.Name.Should().Be("John");
        resultList[0].ObjectResult.IsMember.Should().BeTrue();

        resultList[1].ObjectResult.Age.Should().Be(25);
        resultList[1].ObjectResult.Name.Should().Be("Jane");
        resultList[1].ObjectResult.IsMember.Should().BeFalse();
    }

    /// <summary> Test that batch-reading includes extra properties that are null. </summary>
    [Fact]
    public void BatchReadExcel_NullProperties_ShouldBePresent()
    {
        // Arrange
        var excelDataReader = new ExcelDataReader<FakeEntity>();
        var excelStream = GetTestSpreadsheet();

        // Act
        var result = excelDataReader.BatchReadExcel(excelStream, batchSize: 10).ToList();

        // Assert
        var resultList = result[0].ToList();
        resultList[1].ObjectResult.Name.Should().Be("Jane");
        resultList[1].ExtraProperties.ContainsKey("Attribute-Foo");
        resultList[1].ExtraProperties["Attribute-Foo"].Should().BeNull();
    }

    /// <summary> Test that alphabetical column is correctly calculated to base-26 zero-based index. </summary>
    /// <param name="cellReference">Alpha-numeric value as used for Excel cell references.</param>
    /// <param name="expectedColumnIndex">Expected zero-based index offset for the column.</param>
    [Theory]
    [InlineData("A99", 0)]
    [InlineData("Z67", 25)]
    [InlineData("AA1220", 26)]
    [InlineData("ZZ100", 701)]
    [InlineData("AAA27", 702)]
    [InlineData("ZZZ1622", 18277)]
    [InlineData("ZZZZ1", 475253)]
    [InlineData("ZZZZZ1", 12356629)]
    public void ColumnIndexFromCellReference_ReturnsCorrectValue(string cellReference, int expectedColumnIndex)
    {
        // Arrange
        var excelDataReader = new ExcelDataReader<FakeEntity>();

        // Act & Assert
        ColumnIndexFromCellReference(excelDataReader, cellReference)
            .Should().Be(expectedColumnIndex);
    }

    [UnsafeAccessor(UnsafeAccessorKind.Method, Name = "ColumnIndexFromCellReference")]
    private static extern int ColumnIndexFromCellReference(ExcelDataReader<FakeEntity> @this, string cellReference);

    private static Stream GetTestSpreadsheet() =>
        SpreadsheetGenerator.GetStream(
        [
            [
                nameof(FakeEntity.Id),
                nameof(FakeEntity.IsMember),
                nameof(FakeEntity.Age),
                nameof(FakeEntity.Name),
                nameof(FakeEntity.CreatedDate),
                nameof(FakeEntity.CreatedBy),
                "Attribute-Foo",
            ],
            [1, true, 30, "John", CurrentDateMinusMinutes(30), "John", "Bar"],
            [2, false, 25, "Jane", CurrentDateMinusMinutes(30), "Jane", null],
            [3, true, 40, "Alice", CurrentDateMinusMinutes(30), "Alice", "Bar"],
            [4, false, 22, "Bob", CurrentDateMinusMinutes(30), "Bob", "Bar"],
            [5, false, 1, string.Empty, CurrentDateMinusMinutes(30), string.Empty, "Bar"],
            [6, true, 32, "Sue", CurrentDateMinusMinutes(30), "Sue", "Bar"],
        ]);

    private static DateTimeOffset CurrentDateMinusMinutes(int mins) => DateTimeOffset.Now.AddMinutes(-mins);
}
