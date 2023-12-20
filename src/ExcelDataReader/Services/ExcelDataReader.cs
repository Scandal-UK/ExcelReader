// <copyright file="ExcelDataReader.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReader.Services;

using System.Globalization;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader.Helpers;
using ExcelDataReader.Interfaces;
using ExcelDataReader.Models;

/// <summary> Streams the content of a spreadsheet as a collection of the specified class type. </summary>
/// <typeparam name="T">Class representing the intended structure for returning.</typeparam>
public partial class ExcelDataReader<T> : IExcelDataReader<T>
    where T : class
{
    private readonly Dictionary<uint, string> dateFormatDictionary = ExcelDateFormats();

    /// <inheritdoc/>
    public IEnumerable<IEnumerable<DictionaryMapperResult<T>>> BatchReadExcel(Stream excelStream, int batchSize)
    {
        using SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelStream, false);
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
        var sheetData = workbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

        using var reader = OpenXmlReader.Create(sheetData);
        List<string> columnHeaders = this.ReadColumnHeadersFromFirstRow(reader, workbookPart).ToList();
        foreach (var batch in this.ReadDataBatches(reader, workbookPart, columnHeaders, batchSize))
        {
            yield return batch.Select(data =>
                DictionaryMapper.GenerateObjectMappingFunction<T>(columnHeaders)(data));
        }
    }

    [GeneratedRegex(@"[\d]")] private static partial Regex MatchNumbersRegex();

    /// <remarks>https://msdn.microsoft.com/en-GB/library/documentformat for more info.</remarks>
    private static Dictionary<uint, string> ExcelDateFormats() => new()
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

    private static Row ReadNextRow(OpenXmlReader reader)
    {
        while (reader.Read() && reader.ElementType != typeof(Row))
        {
        }

        return reader.ElementType == typeof(Row) ? (Row)reader.LoadCurrentElement() : null;
    }

    private IEnumerable<List<Dictionary<string, string>>> ReadDataBatches(OpenXmlReader reader, WorkbookPart workbookPart, List<string> columnHeaders, int batchSize)
    {
        var batch = new List<Dictionary<string, string>>();
        int rowCount = 1;

        while (reader.Read())
        {
            if (reader.ElementType == typeof(Row))
            {
                var row = (Row)reader.LoadCurrentElement();
                var rowDictionary = this.ReadRowAsDictionary(row, workbookPart, columnHeaders);
                columnHeaders.ForEach(column => rowDictionary[column] = rowDictionary.TryGetValue(column, out string value) ? value : null);
                batch.Add(rowDictionary);

                if (rowCount % batchSize == 0)
                {
                    yield return batch;
                    batch = [];
                }

                rowCount++;
            }
        }

        if (batch.Count > 0)
        {
            yield return batch;
        }
    }

    private IEnumerable<string> ReadColumnHeadersFromFirstRow(OpenXmlReader reader, WorkbookPart workbookPart) =>
        ReadNextRow(reader)?.Elements<Cell>()
            .Select(cell => this.GetCellValue(cell, workbookPart));

    private string GetCellValue(Cell cell, WorkbookPart workbookPart)
    {
        if (cell.DataType?.Value == CellValues.SharedString)
        {
            var sharedStringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
            if (sharedStringTablePart != null)
            {
                return sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>()
                    .ElementAt(int.Parse(cell.InnerText)).Text.Text;
            }
        }

        if (cell.StyleIndex != null &&
            workbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ChildElements[int.Parse(cell.StyleIndex.InnerText)] is CellFormat cellFormat)
        {
            var dateFormat = this.GetDateTimeFormat(cellFormat.NumberFormatId);
            if (!string.IsNullOrEmpty(dateFormat) && !string.IsNullOrEmpty(cell.InnerText) && double.TryParse(cell.InnerText, out var cellDouble))
            {
                return DateTime.FromOADate(cellDouble).ToString("u", CultureInfo.InvariantCulture);
            }
        }

        return cell.InnerText;
    }

    private string GetDateTimeFormat(UInt32Value numberFormatId) =>
        this.dateFormatDictionary.ContainsKey(numberFormatId) ? this.dateFormatDictionary[numberFormatId] : null;

    private int ColumnIndexFromCellReference(string cellReference) =>
        MatchNumbersRegex().Replace(cellReference, string.Empty)
            .Reverse()
            .Select((c, i) => (c - 'A' + 1) * (int)Math.Pow(26, i))
            .Sum() - 1;

    private Dictionary<string, string> ReadRowAsDictionary(Row row, WorkbookPart workbookPart, List<string> columnHeaders) =>
        row.Elements<Cell>()
            .Where(cell => cell.CellReference != null)
            .Select(cell => new KeyValuePair<string, string>(columnHeaders[this.ColumnIndexFromCellReference(cell.CellReference)], this.GetCellValue(cell, workbookPart)))
            .ToDictionary(kv => kv.Key, kv => kv.Value);
}
