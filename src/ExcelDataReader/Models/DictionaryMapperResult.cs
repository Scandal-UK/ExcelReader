// <copyright file="DictionaryMapperResult.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReader.Models;

/// <summary> Representation of the parsed result of a spreadsheet row. </summary>
/// <typeparam name="T">Object expected from the spreadsheet row.</typeparam>
public class DictionaryMapperResult<T>
{
    /// <summary> Gets or sets the object parsed from a spreadsheet row. </summary>
    public T ObjectResult { get; set; }

    /// <summary> Gets or sets the properties from the record that do not belong to the object. </summary>
    public Dictionary<string, string> ExtraProperties { get; set; }

    /// <summary> Gets or sets the format warnings for the row. </summary>
    public List<string> Warnings { get; set; } = [];
}
