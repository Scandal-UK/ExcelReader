// <copyright file="SpreadsheetGeneratorStubTests.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReaderTests.Helpers;

using System.Collections.Generic;
using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReaderTests.Helpers.Stubs;

/// <summary> Tests the <see cref="SpreadsheetGenerator"/> test-stub class. </summary>
public class SpreadsheetGeneratorStubTests
{
    private static Dictionary<uint, string> DateFormatDictionary => new()
    {
        [14] = "dd/MM/yyyy",
        [15] = "d-MMM-yy",
        [16] = "d-MMM",
        [17] = "MMM-yy",
        [18] = "h:mm AM/PM",
        [19] = "h:mm:ss AM/PM",
        [20] = "h:mm",
        [21] = "h:mm:ss",
        [22] = "M/d/yy h:mm",
        [30] = "M/d/yy",
        [34] = "yyyy-MM-dd",
        [45] = "mm:ss",
        [46] = "[h]:mm:ss",
        [47] = "mmss.0",
        [51] = "MM-dd",
        [52] = "yyyy-MM-dd",
        [53] = "yyyy-MM-dd",
        [55] = "yyyy-MM-dd",
        [56] = "yyyy-MM-dd",
        [58] = "MM-dd",
        [165] = "M/d/yy",
        [166] = "dd MMMM yyyy",
        [167] = "dd/MM/yyyy",
        [168] = "dd/MM/yy",
        [169] = "d.M.yy",
        [170] = "yyyy-MM-dd",
        [171] = "dd MMMM yyyy",
        [172] = "d MMMM yyyy",
        [173] = "M/d",
        [174] = "M/d/yy",
        [175] = "MM/dd/yy",
        [176] = "d-MMM",
        [177] = "d-MMM-yy",
        [178] = "dd-MMM-yy",
        [179] = "MMM-yy",
        [180] = "MMMM-yy",
        [181] = "MMMM d, yyyy",
        [182] = "M/d/yy hh:mm t",
        [183] = "M/d/y HH:mm",
        [184] = "MMM",
        [185] = "MMM-dd",
        [186] = "M/d/yyyy",
        [187] = "d-MMM-yyyy",
    };

    /// <summary> Check that generated spreadsheet stream opens, parses and contains expected data. </summary>
    [Fact]
    public void GeneratedSpreadsheet_ShouldHaveExpectedContent()
    {
        // Arrange
        var testData = GetTestData();
        var generatedStream = SpreadsheetGenerator.GetStream(testData);

        // Act
        using var spreadsheetDocument = SpreadsheetDocument.Open(generatedStream, false);
        var workbookPart = spreadsheetDocument.WorkbookPart;
        var worksheetPart = workbookPart.WorksheetParts.First();
        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

        // Assert
        sheetData.Elements<Row>().Should().HaveCount(testData.Count);

        for (int i = 0; i < testData.Count; i++)
        {
            var rowData = testData[i];
            var row = sheetData.Elements<Row>().ElementAt(i);

            row.Elements<Cell>().Should().HaveCount(rowData.Count);

            for (int j = 0; j < rowData.Count; j++)
            {
                var cell = row.Elements<Cell>().ElementAt(j);
                var actualCellValue = GetCellValue(cell, workbookPart);

                actualCellValue.Should().Be(ObjectToString(rowData[j]));
            }
        }
    }

    private static string ObjectToString(object data)
    {
        if (data is DateTime date)
        {
            return date.ToString("u", CultureInfo.InvariantCulture);
        }
        else if (data is DateTimeOffset dateOffset)
        {
            return dateOffset.ToString("u", CultureInfo.InvariantCulture);
        }

        return data.ToString();
    }

    private static List<List<object>> GetTestData() =>
        new()
        {
            new List<object> { "Header1", "Header2", "Header3", "Header4", "Header5" },
            new List<object> { "Value1", "Value2", "Value3", true, DateTimeOffset.UtcNow.AddDays(-7).ToString("u", CultureInfo.InvariantCulture) },
            new List<object> { "Value4", "Value5", "Value6", false, DateTimeOffset.UtcNow.AddDays(-1).ToString("u", CultureInfo.InvariantCulture) },
        };

    private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType != null)
        {
            if (cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringTablePart != null)
                {
                    var sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(cell.InnerText));
                    return sharedStringItem.Text.Text;
                }
            }
        }

        if (cell.StyleIndex != null)
        {
            if (workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[
                int.Parse(cell.StyleIndex.InnerText)] is CellFormat cellFormat)
            {
                var dateFormat = GetDateTimeFormat(cellFormat.NumberFormatId);
                if (!string.IsNullOrEmpty(dateFormat) && !string.IsNullOrEmpty(cell.InnerText) && double.TryParse(cell.InnerText, out var cellDouble))
                {
                    return DateTime.FromOADate(cellDouble).ToString("u", CultureInfo.InvariantCulture);
                }
            }
        }

        return cell.InnerText;
    }

    private static string GetDateTimeFormat(UInt32Value numberFormatId) =>
        DateFormatDictionary.ContainsKey(numberFormatId) ? DateFormatDictionary[numberFormatId] : null;
}
