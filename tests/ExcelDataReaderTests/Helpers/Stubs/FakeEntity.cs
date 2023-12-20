// <copyright file="FakeEntity.cs" company="Dan Ware">
// Copyright (c) Dan Ware. All rights reserved.
// </copyright>

namespace ExcelDataReaderTests.Helpers.Stubs;

/// <summary> Fake version of an entity to be used as a test stub. </summary>
public class FakeEntity
{
    /// <summary> Gets or sets the ID. </summary>
    public string Id { get; set; }

    /// <summary> Gets or sets the name. </summary>
    public string Name { get; set; }

    /// <summary> Gets or sets the age. </summary>
    public int Age { get; set; }

    /// <summary> Gets or sets a value indicating whether entity is member. </summary>
    public bool IsMember { get; set; }

    /// <summary> Gets or sets the tags. </summary>
    public List<string> Tags { get; set; } = [];

    /// <summary> Gets or sets the numbers. </summary>
    public List<int> Numbers { get; set; } = [];

    /// <summary> Gets or sets the fake entities. </summary>
    public List<FakeEntity> RelatedFakeEntities { get; set; } = [];

    /// <summary> Gets or sets the related entity. </summary>
    public FakeEntity RelatedEntity { get; set; }

    /// <summary> Gets or sets the unmappable. </summary>
    public Dictionary<int, string> Unmappable { get; set; } = [];

    /// <summary> Gets or sets the unparsable. </summary>
    /// <remarks>Although techinically parsable, <see cref="TimeSpan"/> is not in our allowed list.</remarks>
    public TimeSpan Unparsable { get; set; } = new();

    /// <summary> Gets or sets the nullable unparsable. </summary>
    /// <remarks>Although techinically parsable, <see cref="TimeSpan"/> is not in our allowed list.</remarks>
    public TimeSpan? NullableUnparsable { get; set; } = new();

    /// <summary> Gets or sets the created date. </summary>
    public DateTimeOffset? CreatedDate { get; set; }

    /// <summary> Gets or sets the created by. </summary>
    public string CreatedBy { get; set; }
}
