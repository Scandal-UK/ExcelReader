# ExcelDataReader

A simple library designed to stream-read large Excel files and return batches of strongly-typed class records.

Handles columns in any order but assumes the first row is made of column headers.

Returns unmapped columns in a collection of "Extra Properties", useful for additional processing.
