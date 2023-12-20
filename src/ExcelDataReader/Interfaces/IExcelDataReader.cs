// <copyright file="IExcelDataReader.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReader.Interfaces;

using ExcelDataReader.Models;

/// <summary> Streams the content of a spreadsheet to a class tyoe for batch saving. </summary>
/// <typeparam name="T">Class representing the intended structure for saving.</typeparam>
public interface IExcelDataReader<T>
    where T : class
{
    /// <summary> Reads a list of objects parsed from Excel file stream, yielding batches of a specified size. </summary>
    /// <param name="excelStream">Opened stream of the Excel spreadsheet file.</param>
    /// <param name="batchSize">Number of results to yield at a time from the spreadsheet.</param>
    /// <returns>List of <see cref="DictionaryMapperResult{T}"/> each containing the populated object and extra properties as a dictionary.</returns>
    public IEnumerable<IEnumerable<DictionaryMapperResult<T>>> BatchReadExcel(Stream excelStream, int batchSize);
}
