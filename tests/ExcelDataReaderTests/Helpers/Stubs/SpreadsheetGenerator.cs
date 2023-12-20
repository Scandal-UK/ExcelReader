// <copyright file="SpreadsheetGenerator.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReaderTests.Helpers.Stubs;

using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

/// <summary> Stub class to generate a binary stream of a spreadsheet for unit tests. </summary>
internal class SpreadsheetGenerator
{
    /// <summary> Generates a spreadsheet from the given data and returns it as a <see cref="Stream"/>. </summary>
    /// <param name="data">Data to populate the spreadsheet.</param>
    /// <returns>Spreadsheet file content.</returns>
    public static Stream GetStream(List<List<object>> data)
    {
        var memoryStream = new MemoryStream();

        using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
        {
            var sheetData = CreateSheetData(spreadsheetDocument);

            for (int i = 0; i < data.Count; i++)
            {
                sheetData.AppendChild(GetRowFromObjectList(data[i], i + 1));
            }
        }

        memoryStream.Position = 0;
        return memoryStream;
    }

    private static SheetData CreateSheetData(SpreadsheetDocument spreadsheetDocument)
    {
        var workbookPart = spreadsheetDocument.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData());

        var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet() { Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" });

        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = new Stylesheet
        {
            Fonts = new Fonts(new Font()),
            Fills = new Fills(new Fill()),
            Borders = new Borders(new Border()),
            CellStyleFormats = new CellStyleFormats(new CellFormat()),
            CellFormats =
                new CellFormats(
                    new CellFormat(),
                    new CellFormat
                    {
                        NumberFormatId = 22,
                        ApplyNumberFormat = true,
                    }),
        };

        return worksheetPart.Worksheet.GetFirstChild<SheetData>();
    }

    private static Row GetRowFromObjectList(List<object> data, int rowIndex)
    {
        var row = new Row();
        for (int index = 0; index < data.Count; index++)
        {
            if (data[index] != null)
            {
                row.AppendChild(CreateCell(data[index].ToString(), GetCellTypeFromObject(data[index]), $"{CellReferenceLetter(index)}{rowIndex}"));
            }
        }

        return row;
    }

    private static string CellReferenceLetter(int index) => index switch
    {
        0 => "A",
        1 => "B",
        2 => "C",
        3 => "D",
        4 => "E",
        5 => "F",
        6 => "G",
        7 => "H",
        _ => null,
    };

    private static CellValues GetCellTypeFromObject(object cellValue) => cellValue switch
    {
        bool => CellValues.Boolean,
        DateTime or DateTimeOffset => CellValues.Date,
        int or decimal or double or long or short or float => CellValues.Number,
        _ => CellValues.String,
    };

    private static Cell CreateCell(string content, CellValues dataType, string cellRef) =>
        (dataType == CellValues.Date) ?
        new(new CellValue(DateTime.Parse(content).ToOADate().ToString(CultureInfo.InvariantCulture))) { CellReference = cellRef, StyleIndex = 1 } :
        new(new CellValue(content)) { DataType = new EnumValue<CellValues>(dataType), CellReference = cellRef };
}
