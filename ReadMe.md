# ExcelDataReader

## Description

A simple library designed to stream-read large Excel files and return batches of strongly-typed class records.

Uses the Simple API for XML (SAX) approach to parse a stream and maps each row to a specified class (ideal for reading from Azure Blob storage). Records are returned in batches of a specified size for efficient processing.

Handles columns in any order but assumes the first row is made up of the column headers.

Returns unmapped columns in a collection of "Extra Properties", useful for additional processing.

## Usage

Please see the unit tests for examples of usage. Documentation will come later.

### Note that this library will likely be renamed before it gets packaged!

(Work-in-progress - intended to be published to NuGet.)
