# LargeXlsx - A .net library to write large XLSX files

[![NuGet](https://img.shields.io/nuget/v/LargeXlsx.svg)](https://www.nuget.org/packages/LargeXlsx)

This is a minimalistic yet feature-rich library, written in C# targeting .NET Standard 2.0 and .NET Core 3.1 or later, providing simple primitives to write Excel files in XLSX format in a streamed manner, so that potentially huge files can be created while consuming a low, constant amount of memory.

Documentation for the old major version 1.x can be found [here](https://github.com/salvois/LargeXlsx/tree/release/v1).


## Supported features

Currently the library supports:

* cells containing inline or shared strings; numeric values; date and time; formulas
* multiple worksheets
* merged cells
* split panes, a.k.a. frozen rows and columns
* styling such as font face, size, and color; background color; cell borders; numeric formatting; alignment
* column and row formatting (custom width, height, hidden columns/rows)
* auto filter
* cell validation, such as dropdown list of allowed values
* right-to-left worksheets, to support languages such as Arabic and Hebrew
* password protection of sheets against accidental modification
* headers, footers and page breaks for worksheet printout


## Example

To create a basic single-sheet Excel document:

```csharp
using var stream = new FileStream("Basic.xlsx", FileMode.Create, FileAccess.Write);
using var xlsxWriter = new XlsxWriter(stream);
xlsxWriter
    .BeginWorksheet("Sheet 1")
    .BeginRow().Write("Name").Write("Location").Write("Height (m)")
    .BeginRow().Write("Kingda Ka").Write("Six Flags Great Adventure").Write(139)
    .BeginRow().Write("Top Thrill Dragster").Write("Cedar Point").Write(130)
    .BeginRow().Write("Superman: Escape from Krypton").Write("Six Flags Magic Mountain").Write(126);
```

To create an Excel document with some fancy formatting:

```csharp
using var stream = new FileStream("Simple.xlsx", FileMode.Create, FileAccess.Write);
using var xlsxWriter = new XlsxWriter(stream);
var headerStyle = new XlsxStyle(
    new XlsxFont("Segoe UI", 9, Color.White, bold: true),
    new XlsxFill(Color.FromArgb(0, 0x45, 0x86)),
    XlsxStyle.Default.Border,
    XlsxStyle.Default.NumberFormat,
    XlsxAlignment.Default);
var highlightStyle = XlsxStyle.Default.With(new XlsxFill(Color.FromArgb(0xff, 0xff, 0x88)));
var dateStyle = XlsxStyle.Default.With(XlsxNumberFormat.ShortDateTime);
var borderedStyle = highlightStyle.With(
    XlsxBorder.Around(new XlsxBorder.Line(Color.DeepPink, XlsxBorder.Style.Dashed)));

xlsxWriter
    .BeginWorksheet("Sheet 1",
        columns: new[] { XlsxColumn.Unformatted(count: 2), XlsxColumn.Formatted(width: 20) })
    .SetDefaultStyle(headerStyle)
    .BeginRow().AddMergedCell(2, 1).Write("Col1").Write("Top2").Write("Top3")
    .BeginRow().Write().Write("Col2").Write("Col3")
    .SetDefaultStyle(XlsxStyle.Default)
    .BeginRow().Write("Row3").Write(42).WriteFormula(
        $"{xlsxWriter.GetRelativeColumnName(-1)}{xlsxWriter.CurrentRowNumber}*10", highlightStyle)
    .BeginRow().Write("Row4").SkipColumns(1).Write(new DateTime(2020, 5, 6, 18, 27, 0), dateStyle)
    .SkipRows(2)
    .BeginRow().Write("Row7", borderedStyle, columnSpan: 2).Write(3.14159265359)
    .SetAutoFilter(2, 1, xlsxWriter.CurrentRowNumber - 1, 3);
```

The output is like:

![Single sheet Excel document with 7 rows and 3 columns](https://github.com/salvois/LargeXlsx/raw/master/example.png)



## Table of contents

  - [Supported features](#supported-features)
  - [Example](#example)
  - [Changelog](#changelog)
  - [Migrating from version 1.x to 2.x](#migrating-from-version-1x-to-2x)
    - [Breaking changes](#breaking-changes)
    - [New features](#new-features)
  - [Usage](#usage)
    - [The insertion point](#the-insertion-point)
    - [Column names](#column-names)
    - [Buffered writes](#buffered-writes)
    - [Creating a new worksheet](#creating-a-new-worksheet)
      - [Column formatting](#column-formatting)
      - [Auto fit column width](#auto-fit-column-width)
    - [Adding or skipping rows](#adding-or-skipping-rows)
    - [Writing cells](#writing-cells)
      - [Writing hyperlinks](#writing-hyperlinks)
      - [Skipping columns](#skipping-columns)
    - [Merged cells](#merged-cells)
    - [Auto filter](#auto-filter)
    - [Data validation](#data-validation)
    - [Password protection of sheets](#password-protection-of-sheets)
    - [Headers and footers](#headers-and-footers)
      - [Left, center and right sections of header and footer](#left-center-and-right-sections-of-header-and-footer)
      - [Header and footer formatting codes](#header-and-footer-formatting-codes)
    - [Page breaks](#page-breaks)
    - [Styling](#styling)
      - [The default style](#the-default-style)
      - [Fonts](#fonts)
      - [Fills](#fills)
      - [Borders](#borders)
      - [Number formats](#number-formats)
      - [Alignment](#alignment)
  - [Special thanks](#special-thanks)
  - [License](#license)


## Changelog

* 2.0: No external dependencies on .NET Core 3.1 or greater, significant performance improvements, async API
* 1.12: Page breaks for printing, inspired by [Nael Al Abbasi](https://github.com/thedude61636)
* 1.11: Validation and optional skipping of invalid XML characters, inspired by [Anton Mihai](https://github.com/mike101200)
* 1.10: Ability to hide worksheets, thanks to [Micha Voße](https://github.com/piwonesien)
* 1.9: Optionally force writing cell references for compatibility with some readers, thanks to [Mikk182](https://github.com/Mikk182); header and footer functionality thanks to [soend](https://github.com/soend)
* 1.8: Ability to hide grid lines and row and column headers from worksheets, thanks to [Rajeev Datta](https://github.com/rajeevdatta)
* 1.7: Write overload for booleans, more performance improvements thanks to [Antony Corbett](https://github.com/AntonyCorbett) and [Mark Pflug](https://github.com/MarkPflug)
* 1.6: Opt-in shared string table (for memory vs. file size trade-off)
* 1.5: Password protection of sheets
* 1.4: Support for font underline, thanks to [Sergey Bochkin](https://github.com/sbochkin)
* 1.3: Optional ZIP64 support for really huge files
* 1.2: Right-to-left worksheets. Parametric compression level to trade between space and speed
* 1.1: Number format for text-formatted cells (the "@" formatting)
* 1.0: Finalized API
* 0.0.9: Started writing XLSX files directly rather than using Microsoft's [Office Open XML library](https://github.com/OfficeDev/Open-XML-SDK)


## Migrating from version 1.x to 2.x

Major version 2 reworked the way XLSX file content is written, by using native UTF-8 strings, and reducing memory allocations and memory copies as much as possible. This allowed for **significant performance improvements**.

Moreover, the SharpCompress compression library has been replaced by System.IO.Compression functionality from the .NET runtime, which provided further performance benefit and removed external dependencies as well. Unfortunately, System.IO.Compression produces valid ZIP64 files (necessary for very big files) only on .NET Core 3.1 or greater, thus, **LargeXlsx is now multi-target** and SharpCompress is still used on the .NET Standard 2.0 target (basically, for compatibility with .NET Framework). Also, some functionality to avoid memory copies is not available on .NET Standard 2.0, hence performance is expected to be not as good.

Thanks to the new way to write XLSX file content, an **asynchronous API** has been introduced, useful, for example, when you want to write XLSX files in ASP.NET Core (Kestrel) scenarios. A dynamically-sized, internal buffer accumulates writes, that are periodically committed to the underlying stream the XLSX file is written to. This is a compromise (suggested by [Antony Corbett](https://github.com/AntonyCorbett)) to let I/O work asynchronously while reducing breaking changes and performance penalty of async state management.

###  Breaking changes

* On .NET Core 3.1 or greater, the SharpCompress library is no longer used. If you relied on SharpCompress as a transitive dependency, you must reference it on your own. SharpCompress is still used on the .NET Standard 2.0 target
* The constructor of the `XlsxWriter` class changed the type of the  `compressionLevel` parameter to a custom enum that no longer depends on the SharpCompress library. If you explicitly specified a compression level, you must change the argument to a sensible equivalent
* The constructor of the `XlsxWriter` class has no `useZip64` parameter any longer, and ZIP64 compression is enabled by default for very big files. If you specified an argument you must remove it

### New features

* The `XlsxWriter` class now provides methods (`Commit`, `TryCommit`, `CommitAsync`, `TryCommitAsync`) to commit its internal buffer to the stream the XLSX file is being written to. The `commitThreshold` parameter of the `XlsxWriter` constructor sets the size in bytes (56 KiB by default, to fit a 64 KiB buffer easily) that trigger actual writes to the underlying stream by the `TryCommit` methods
* `BeginWorksheet` and `BeginRow` also try to commit writes to the underlying stream implicitly, hence you are not forced to use the commit methods above. Moreover, they gained `Async` overloads
* `XlsxWriter`, which already used to be `IDisposable`, is now also `IAsyncDisposable`, so that it can be used with `await using` to let final writes be committed asynchronously


## Usage

The `XlsxWriter` class is the entry point for almost all functionality of the library. It is designed so that most of its methods can be chained to write the Excel file using a fluent syntax.

The constructor allows you to create an XLSX writer. Please note that an `XlsxWriter` object **must be disposed**, either with `Dispose` or `DisposeAsync`, to properly finalize the Excel file. Sandwiching its lifetime in a `using` or `async using` statement is recommended. Moreover, an `XlsxWriter` is not designed to be used on multiple threads.

```csharp
// class XlsxWriter
public XlsxWriter(
    Stream stream,
    XlsxCompressionLevel compressionLevel = XlsxCompressionLevel.Fastest,
    bool requireCellReferences = true,
    bool skipInvalidCharacters = false,
    int commitThreshold = 57344);
```

The constructor accepts:
* A writeable `Stream` to save the Excel file into
* An optional desired compression level of the underlying zip stream. The default `XlsxCompressionLevel.Fastest` roughly matches file sizes produced by Excel while providing good performance. The alternative `XlsxCompressionLevel.Optimal` enum value provides higher compression but may result in lower speed
* An optional flag indicating whether row numbers and cell references (such as "A1") are to be included in the XLSX file even when redundant. Row numbers and cell references are optional according to the specification, and omitting them provides a notable performance boost when writing XLSX files (as much as 40%). Unfortunately, some non-compliant readers (which apparently [include MS Access itself](https://github.com/salvois/LargeXlsx/issues/36)!) consider files without row and cell references as invalid, thus you can be conservative and set this flag to `true` if you want to make them happy. Spreadsheet applications such as Excel and LibreOffice can read XLSX files without references just fine, thus, if they are your target, you could use `false` for greater performance
* An optional flag indicating how to behave when trying to write characters that are invalid for the XML underlying the XLSX file format. When `false`, an XmlException is thrown if such invalid characters are found. When `true`, invalid characters are just skipped
* A `commitThreshold` value representing the size in bytes the internal buffer of the `XlsxWriter` must exceed to trigger actual writes to the underlying stream by `TryCommit`, `BeginRow` and their `Async` overloads, to avoid micro-commits

The recipe is adding a worksheet with `BeginWorksheet`, adding a row with `BeginRow`, writing cells to that row with `Write`, and repeating as required. Rows and worksheets are implicitly finalized as soon as new rows or worksheets are added, or the `XlsxWriter` is disposed.


### The insertion point

To enable streamed writes, the content of the Excel file must be written strictly from top to bottom and from left to right. Think of an insertion point always advancing when writing content.\
The `CurrentRowNumber` and `CurrentColumnNumber` read-only properties will return the location of the **next** cell that will be written. Both the row and column numbers are **one-based**.

```csharp
// class XlsxWriter
public int CurrentRowNumber { get; }
public int CurrentColumnNumber { get; }
```

Please note that `CurrentColumnNumber` may be zero, and thus invalid, if the current row has not been set up using `BeginRow` (attempting to write a cell would throw an exception).


### Column names

In the usual "A1" cell reference format, columns are named from A (for column number 1) to XFD (for column number 16384).\
The following facilitate conversion from column numbers to column names, useful when writing formulas:

```csharp
// class XlsxWriter
public string CurrentColumnName { get; }
public string GetRelativeColumnName(int offsetFromCurrentColumn);
public static string GetColumnName(int columnIndex);
```

The first version returns the column name for the column at the insertion point. The second version returns the column name for a column relative to the insertion point. The last version returns the column name for an absolute column index. Absolute or relative indexes outside the range [1..16384] will result in an `ArgumentOutOfRangeException`.

### Buffered writes

The following methods operate on the internal buffer of the `XlsxWriter` object:

```csharp
// class XlsxWriter
public int BufferCapacity { get; }
public XlsxWriter Commit();
public Task<XlsxWriter> CommitAsync();
public XlsxWriter TryCommit();
public Task<XlsxWriter> TryCommitAsync();
```

The `BufferCapacity` property returns the capacity, in bytes, currently allocated to the internal buffer. It is intended as debugging facility, as the buffer is not expected to grow bigger than tens or hundreds of kilobytes in normal usage.

The `Commit` and `CommitAsync` methods write the content of the buffer to the underlying stream the XLSX file is being written to, and empty the buffer. You can use them if you want to force commit and you don't want to wait, for example, for the next `BeginRow`.

The `TryCommit` and `TryCommitAsync` methods do commit only if the content of the buffer exceeds the `commitThreshold` argument passed to the `XlsxWriter` constructor. Thus, you can call them when you see fit, without worrying about micro-commits. `BeginRow` and `BeginRowAsync` also do the same, hence you don't need to commit explicitly in a typical use case.


### Creating a new worksheet

Call `BeginWorksheet` or `BeginWorksheetAsync` passing the sheet name and one or more of the following optional parameters (using named arguments is recommended):

- `splitRow`: if greater than zero, the one-based index of the row where to place a horizontal split to create frozen panes
- `splitColumn`: if greater than zero, the one-based index of the column where to place a vertical split to create frozen panes
- `rightToLeft`: set to `true` to switch the worksheet to right-to-left mode (to support languages such as Arabic and Hebrew)
- `columns`: pass a non-`null` list to specify optional column formatting (see below)
- `showGridLines`: set to `false` to hide gridlines in the sheet
- `showHeaders`: set to `false` to hide row and column headers in the sheet
- `state`: whether the worksheet is `Visible`, `Hidden` (but spreadsheet applications let you unhide it) or `VeryHidden` (spreadsheet applications are not supposed to let you unhide it; Excel doesn't but LibreOffice does)

```csharp
// class XlsxWriter
public XlsxWriter BeginWorksheet(
    string name,
    int splitRow = 0,
    int splitColumn = 0,
    bool rightToLeft = false,
    IEnumerable<XlsxColumn> columns = null,
    bool showGridLines = true,
    bool showHeaders = true,
    XlsxWorksheetState state = XlsxWorksheetState.Visible);
public Task<XlsxWriter> BeginWorksheetAsync(...);
```

Note that, for compatibility with a limitation of the Excel application, names are restricted to a maximum of 31 characters. An `ArgumentException` is thrown if a longer name is passed.
An `ArgumentException` is also thrown when trying to add a worksheet with a name already used for another worksheet.

A call to `BeginWorksheet` or `BeginWorksheetAsync` finalizes the last worksheet being written, if any, and sets up a new one, so that rows can be added.


#### Column formatting

The `BeginWorksheet` and `BeginWorksheetAsync` methods accept an optional `columns` parameter to specify a list of column formatting objects of type `XlsxColumn`, each describing one or more adjacent columns, with their custom width, hidden state or default style, starting from column A.

This information must be provided before writing any content to the worksheet, thus the number, width and styles of the columns must be known in advance.

You can create `XlsxColumn` objects with one of these named constructors:

```csharp
// class XlsxColumn
public static XlsxColumn Unformatted(int count = 1);
public static XlsxColumn Formatted(double width, int count = 1, bool hidden = false, XlsxStyle style = null);
```

`Unformatted` creates a column description that is used basically to skip one or more unformatted columns.

`Formatted` creates a column description to specify the mandatory width, optional hidden state, and optional style of one or more contiguous columns. The width is expressed (simplifying) in approximate number of characters. The column style represents how to style all *empty* cells of a column. Cells that are explicitly written always use the cell style instead.


#### Auto fit column width

A frequently asked question about column formatting is how to automatically set column width based on column content.

Unfortunately, the XLSX file format does not provide any special feature for automatic/best fit column widths, so your best bet is to estimate your maximum or average content width and use it in column formatting as described above. Moreover, since the file format specifies that column formatting is written into the file before cell contents, you are required to estimate width before streaming any data, or iterating your dataset twice, or just guessing a sensible value (the latter usually works surprisingly well).

To be fair, the XLSX specification provides a `bestFit` attribute for columns, but it has a totally different purpose (automatically enlarging a column when a user types digits in it).


### Adding or skipping rows

Call `BeginRow` or `BeginRowAsync` to advance the insertion point to the beginning of the next line and set up a new row to accept content. If a previous row was being written, it is finalized before creating the new one.

```csharp
// class XlsxWriter
public XlsxWriter BeginRow(double? height = null, bool hidden = false, XlsxStyle style = null);
public Task<XlsxWriter> BeginRowAsync(...);
```
    
You can specify optional row formatting when creating a new row. The height is expressed in points. The row style represents how to style all *empty* cells of a row. Cells that are explicitly written always use the cell style instead.

Calling `BeginRow` or `BeginRowAsync` also acts as if you called `TryCommit` or `TryCommitAsync`, respectively, hence any buffered content is written to the underlying stream the XLSX file is being written to, if the buffered content exceeds the `commitThreshold` argument passed to the `XlsxWriter` constructor. This is to keep writing content practical and to avoid big breaking changes from major version 1, while being explicit with moments where actual writes may occur.

Call `SkipRows` to move the insertion point down by the specified count of rows, that will be left empty and unstyled (unless column styles are in place). If a previous row was being written, it is finalized. Please note that `BeginRow` must be called anyway before starting to write a new row.

```csharp
// class XlsxWriter
public XlsxWriter SkipRows(int rowCount);
```

### Writing cells

Call one of the `Write` methods to write content to the cell at the insertion point. Note that content is written to the internal buffer and not directly to the underlying stream the XLSX file is being written to, therefore they are all synchronous operations.

```csharp
// class XlsxWriter
public XlsxWriter Write(XlsxStyle style = null, int columnSpan = 1, int repeatCount = 1);
public XlsxWriter Write(string value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(double value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(decimal value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(int value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(DateTime value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(bool value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter WriteFormula(string formula, XlsxStyle style = null, int columnSpan = 1, IConvertible result = null);
public XlsxWriter WriteSharedString(string value, XlsxStyle style = null, int columnSpan = 1);
```

 You may write one of the following:

  * **Nothing**: a cell containing no value, but styled nonetheless.
  * **String**: an inline literal string of text; if the string is `null` the method falls back on the "nothing" case. The value of an inline string is written into the cell, thus resulting in low memory consumption but possibly larger files (see, in contrast, shared strings).
  * **Number**: a numeric constant, either as an `int`, a `double` or a `decimal`. Note that applications such as Excel and LibreOffice interpret numbers as `double`, so expect precision loss if you write values that cannot be represented exactly by a `double`.
  * **Date and time**: a `DateTime` value, that will be converted to its `double` representation (days since 1900-01-01). Note that you must style the cell using a date/time number format to have the value appear as a date.
  * **Boolean**: a `bool` value, that will appear either as `TRUE` or `FALSE`
  * **Formula**: a string that Excel or a compatible application will interpret as a formula to calculate. Note that, unless you provide a `result` calculated by yourself (either string or numeric), no result is saved into the XLSX file. However, a spreadsheet application will calculate the result as soon as the XLSX file is opened.
  * **Shared string**: a shared literal string of text. Unlike inline strings, shared strings are saved in a look-up table and deduplicated in constant time, and only a reference is written into the cell. This may help to produce smaller files, but, since the shared strings must be accumulated in RAM, **writing a large number of *different* shared strings may cause high memory consumption**. Keep this in mind when choosing between inline strings and shared strings.

The `style` parameter specifies the style to use for the cell being written. If `null` (or omitted), the cell is styled using the current default style of the `XlsxWriter` (see Styling). Note that in no case the style of the column or the row, if any, is used for written cells.

When writing empty cells you can also specify a `repeatCount` parameter, to write multiple consecutive styled empty cells. Note that the difference between `repeatCount` and `columnSpan` both greater than 1 is that the latter creates a merged cell (with its memory consumption drawback) and does not have in-between borders.

The `columnSpan` parameter can be used to let the cell span multiple columns. When greater than 1, a merged range of such cells is created (see Merged cells), content is written to the first cell of the range and the insertion point is advanced after the merged cells. Note that `xlsxWriter.Write(value, columnSpan: count)` is actually a shortcut for `xlsxWriter.AddMergedCells(1, count).Write(value).Write(repeatCount: count - 1)`. Since a merged cell is created, **writing a large number of cells with columnSpan greater than 1 may cause high memory consumption**.

#### Writing hyperlinks

The XLSX file format provides two ways to insert hyperlinks (either to a web site or cells perhaps in a different workbook): using a specific "hyperlinks" section in the worksheet and using the `HYPERLINK` formula in a cell.

The former, which is the one used by the "Insert hyperlink" command in Excel, is not used by this library, because it would require to accumulate all hyperlinks in RAM until the worksheet is complete.

The latter can be used with `WriteFormula` while streaming content into the worksheet, thus is more appropriate for the use case of this library, for example:

```csharp
xlsxWriter.WriteFormula(
    "HYPERLINK(\"https://github.com/salvois/LargeXlsx\", \"LargeXlsx on GitHub\")",
    XlsxStyle.Default.With(XlsxFont.Default.WithUnderline().With(Color.Blue));
```

where the first parameter of the `HYPERLINK` formula is the link location and the second parameter, which is optional, is a friendly name to display into the cell. Styling may be used to show the cell contains a link.

#### Skipping columns

Like rows, cells can be skipped using the `SkipColumns` method, to move the insertion point to the right by the specified count of cells, that will be left empty and unstyled (unless column or row styles are in place).

```csharp
// class XlsxWriter
public XlsxWriter SkipColumns(int columnCount);
```

### Merged cells

A rectangle of adjacent cells can be merged using the `AddMergedCells` method:

```csharp
// class XlsxWriter
public XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount);
public XlsxWriter AddMergedCell(int rowCount, int columnCount);
```

The first overload lets you specify an arbitrary rectangle in the worksheet.\
The second overload facilitates merging cells while fluently writing the file, using the insertion point as the top-left cell for the merged range.

Creating a merged cell range does not advance the insertion point.\
As a shortcut for the common case of merging a 1 row by n columns range, you can pass a value greater than 1 as the `columnSpan` argument of a `Write` method, that will create the merged cell and advance the insertion point for you.

Content for the merged cells must be written in the top-left cell of the rectangle. A spreadsheet application will not display any content of the remaining cells in the merged range. Thus, you should explicitly skip those cells using `SkipColumns`, `SkipRows` and writing empty cells (if styling is needed) as appropriate.\
For example, if merging the 2 rows x 3 columns range `A7:C8` using `AddMergedCells(7, 1, 2, 3)`, you must write content for the merged cell in `A7`, then explicitly jump by further 2 columns using `SkipColumns(2)` to continue writing content from `D7`, and the same applies on row 8, where after a `BeginRow()` you must skip 3 columns with `SkipColumns(3)` and continue writing from `D8`.

**Note**: due to the structure of the XLSX file format, the ranges for all merged cells of a worksheet must be accumulated in RAM, because they must be written to the file after the content of the whole worksheet. **Using a large number of merged cells may cause high memory consumption**. This also means that you may call `AddMergedCell` at any moment while you are writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object), even for cells already written or well before writing them, or cells you won't write content to.


### Auto filter

You can add an auto filter (the one created with the funnel icon in Excel) for a specific rectangular region, containing headers in the first row, using:

```csharp
// class XlsxWriter
public XlsxWriter SetAutoFilter(int fromRow, int fromColumn, int rowCount, int columnCount);
```

You can call `SetAutoFilter` at any moment while writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object). Each worksheet can contain only up to one auto filter, thus if you call `SetAutoFilter` multiple times for the same worksheet only the last one will apply.


### Data validation

Data validation lets you add constraints on cell content. Such constraints are  represented by `XlsxDataValidation` objects, created with:

```csharp
// class XlsxDataValidation
public XlsxDataValidation(
        bool allowBlank = false,
        string error = null,
        string errorTitle = null,
        XlsxDataValidation.ErrorStyle? errorStyle = null,
        XlsxDataValidation.Operator? operatorType = null,
        string prompt = null,
        string promptTitle = null,
        bool showDropDown = false,
        bool showErrorMessage = false,
        bool showInputMessage = false,
        XlsxDataValidation.ValidationType? validationType = null,
        string formula1 = null,
        string formula2 = null);
```

Using named arguments is recommended to improve readability. The parameters represent:
- `allowBlank`: whether an empty cell is considered valid
- `error`: an optional error message to replace the default message of the spreadsheet application when invalid content is detected
- `errorTitle`: an optional title to replace the default title of the spreadsheet application when invalid content is detected
- `errorStyle`: whether to report detection of invalid content as a blocking error, a warning or a notice
- `operatorType`: for validation involving comparison (see `validationType`), the comparison operator to apply
- `prompt`: an optional message to be shown as a tooltip when the cell receives focus
- `promptTitle`: an optional title to show in the prompt tooltip
- `showDropDown`: whether to show a dropdown list when the validation type is set to `XlsxDataValidation.ValidationType.List`; **Note:** for some reason, this seems to work backwards, with both Excel and LibreOffice showing the dropdown when this property is **false**
- `showErrorMessage`: whether the spreadsheet application will show an error when invalid content is detected
- `showInputMessage`: whether the spreadsheet application will show the prompt tooltip when a validated cell is focused
- `validationType`: specifies to do validation on numeric values, whole integer values, date values, text length (using the comparison operator specified by `operatorType`) or against a list of allowed values
- `formula1`: the value to compare the cell content against; for lists, it can be a reference to a cell range containing the allowed values, or a list of comma separated case-sensitive string constants, enclosed by double quotes (e.g. `"\"item1,item2\"")`; for other validation types, it can be a reference to a cell containing the value or a constant
- `formula2`: for comparison involving two values (e.g. operator `Between`), the second value, specified as in `formula1`.

To ease creation of an `XlsxDataValidation` specifying validation against a list of string constants, the following named constructor is provided. Note that it basically does a `string.Join` of choices into `formula1`, thus if a choice includes a comma, the spreadsheet application will split it into separate choices. There is no way to include real commas in choices specified as string constants (rather than as cell range).

```csharp
// class XlsxDataValidation
public static XlsxDataValidation List(
        IEnumerable<string> choices,
        bool allowBlank = false,
        string error = null,
        string errorTitle = null,
        XlsxDataValidation.ErrorStyle? errorStyle = null,
        string prompt = null,
        string promptTitle = null,
        bool showDropDown = false,
        bool showErrorMessage = false,
        bool showInputMessage = false);
```

To add validation rules to a worksheet, use one of the following while writing content:

```csharp
// class XlsxWriter
public XlsxWriter AddDataValidation(int fromRow, int fromColumn, int rowCount, int columnCount,
                                    XlsxDataValidation dataValidation);
public XlsxWriter AddDataValidation(int rowCount, int columnCount, XlsxDataValidation dataValidation);
public XlsxWriter AddDataValidation(XlsxDataValidation dataValidation);
```

The first overload applies the validation rules to all cells in the specified rectangular range. The second overload uses the insertion point as the top-left corner of the rectangular range. The third overload applies validation only on the cell at the insertion point.

Note that, due to the internals of the XLSX file format, all validation objects and their cell references must be kept in RAM until a worksheet is finalized, but this library **deduplicates** validation objects in **constant time** as needed. Thus, you should usually not worry about performance or memory consumption when you use multiple validation objects, unless you are using a large number of different ones, or specify a lot of separate cell references.

This means that you may call `AddDataValidation` at any moment while you are writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object), even for cells already written or well before writing them, or cells you won't write content to.


### Password protection of sheets

Password protection of worksheets helps prevent accidental modification of data. You can enable protection on the worksheet being written using:

```csharp
// class XlsxWriter
public XlsxWriter SetSheetProtection(XlsxSheetProtection sheetProtection);

// class XlsxSheetProtection
public XlsxSheetProtection(
        string password,
        bool sheet = true,
        bool objects = true,
        bool scenarios = true,
        bool formatCells = true,
        bool formatColumns = true,
        bool formatRows = true,
        bool insertColumns = true,
        bool insertRows = true,
        bool insertHyperlinks = true,
        bool deleteColumns = true,
        bool deleteRows = true,
        bool selectLockedCells = false,
        bool sort = true,
        bool autoFilter = true,
        bool pivotTables = true,
        bool selectUnlockedCells = false);
```

Each flag in `XlsxSheetProtection` specifies what operations are **protected**, that is not allowed. By default, only selecting cells is allowed. The password must have a length between 1 and 255 characters.

You can call `SetSheetProtection` at any moment while writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object). Each worksheet can contain only up to one protection definition, thus if you call `SetSheetProtection` multiple times for the same worksheet only the last one will apply.

**Note:** password protection of sheets is not to be confused with workbook encryption and is not meant to be secure. File contents are still written in clear text and may be changed by deliberately editing the file. The password is not written into the file but a hash of the password is. Workbook encryption (that is, password protection of the whole file) requires the writer to use a different file format and will be, unfortunately, not provided by this library in the foreseeable future.

### Headers and footers

Call `SetHeaderFooter` on an `XlsxWriter` when you want to add headers and footers to a worksheet printout:

```csharp
//class XlsxWriter
public XlsxWriter SetHeaderFooter(XlsxHeaderFooter headerFooter);

//class XlsxHeaderFooter
public XlsxHeaderFooter(
        string oddHeader = null,
        string oddFooter = null,
        string evenHeader = null,
        string evenFooter = null,
        string firstHeader = null,
        string firstFooter = null,
        bool alignWithMargins = true,
        bool scaleWithDoc = true);
```

You can call `SetHeaderFooter` at any moment while writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object). Each worksheet can contain only up to one header/footer definition, thus if you call `SetHeaderFooter` multiple times for the same worksheet only the last one will apply.

The constructor of `XlsxHeaderFooter` takes several properties to control the content of headers or footers:
- `oddHeader` sets the header text, if any, to use on every page, or only on odd pages if `evenHeader` is also set
- `oddFooter` sets the footer text, if any, to use on every page, or only on odd pages if `evenFooter` is also set
- `firstHeader` sets the header text, if any, to be used only on the first page
- `firstFooter` sets the footer text, if any, to be used only on the first page
- `evenHeader` sets the header text, if any, to be used only on even pages
- `evenFooter` sets the footer text, if any, to be used only on even pages
- `alignWithMargins`: align header and footer margins with page margins. When true, as left/right margins grow and shrink, the header and footer edges stay aligned with the margins. When false, headers and footers are aligned on the paper edges, regardless of margins
- `scaleWithDoc`: scale header and footer with document scaling

**Note**: that header and footer text may contain a number of formatting codes to control styling and insert special content such as the page number or file name. These codes are prefixed with the `&` character (more information below). A single literal ampersand must thus be represented as `&&`.

#### Left, center and right sections of header and footer

Spreadsheet applications such as Excel and LibreOffice expect headers and footers to comprise three different sections, for left-, center- and right-aligned texts.

These sections are introduced by the `&L`, `&C` and `&R` formatting codes respectively. For maximum compatibility, you should always use at least one of them to specify at least one section, and you should not repeat the same sections multiple times (although supported, this could lead to unexpected results)

Other libraries force you to follow the above recommendation by providing an object with three separate properties rather than a single string. This library does not enforce this by design, to facilitate localization in right-to-left cultures.

#### Header and footer formatting codes

Excel specifies a set of formatting codes to control styling and insert special content into headers and footers, as specified in [Office Implementation Information for ISO/IEC 29500 Standards Support](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/c167a243-45ad-4def-816e-7032fb1adf5c).

You can either insert these formatting codes manually in your header and footer strings, to ease localization, or use the following class to build header and footer strings programmatically:

```csharp
public class XlsxHeaderFooterBuilder
{
    public override string ToString(); // renders the built text

    public XlsxHeaderFooterBuilder Left(); // &L - left section
    public XlsxHeaderFooterBuilder Center(); // &C - center section
    public XlsxHeaderFooterBuilder Right(); // &R - right section
    public XlsxHeaderFooterBuilder Text(string text); // plain text, auto escaping ampersands
    public XlsxHeaderFooterBuilder CurrentDate(); // &D
    public XlsxHeaderFooterBuilder CurrentTime(); // &T
    public XlsxHeaderFooterBuilder FileName(); // &F
    public XlsxHeaderFooterBuilder FilePath(); // &Z
    public XlsxHeaderFooterBuilder NumberOfPages(); // &N
    public XlsxHeaderFooterBuilder PageNumber(int offset = 0) // &P or &P+offset or &P+offset - current page number, optionally offset by the specified number
    public XlsxHeaderFooterBuilder SheetName(); // &A
    public XlsxHeaderFooterBuilder FontSize(int points); // &points
    public XlsxHeaderFooterBuilder Font(string name, bool bold = false, bool italic = false); // &"name,type" - set font name and type
    public XlsxHeaderFooterBuilder Font(bool bold = false, bool italic = false); // &"-,type" - set only font type
    public XlsxHeaderFooterBuilder Bold(); // &B - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder Italic(); // &I - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder Underline(); // &U - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder DoubleUnderline(); // &E - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder StrikeThrough(); // &S - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder Subscript(); // &Y - each occurrence toggles on or off
    public XlsxHeaderFooterBuilder Superscript(); // &X - each occurrence toggles on or off
}
```

When writing the `&"font,type"` and `&"-,type"` codes manually, the type field can assume one of the following values: `Regular`, `Bold`, `Italic` and `Bold Italic`.


### Page breaks

The following methods of `XlsxWriter` can be used to control how to split a worksheet printout either horizontally or vertically:

```csharp
// class XlsxWriter
public XlsxWriter AddRowPageBreakBefore(int rowNumber);
public XlsxWriter AddColumnPageBreakBefore(int columnNumber);
public XlsxWriter AddRowPageBreak();
public XlsxWriter AddColumnPageBreak();
```

The former two methods let you specify the row and column number before which a horizontal or vertical page break is placed, respectively.

The latter two place a horizontal or vertical page break before the current row or column number, respectively (see [The insertion point](#the-insertion-point)), and make sense when writing content fluently.

Trying to place a break before row 1 or column 1 (that is, column A) will result in an `ArgumentOutOfRangeException`.

**Note**: due to the structure of the XLSX file format, page breaks of a worksheet must be accumulated in RAM, because they must be written to the file after the content of the whole worksheet. **Using a very large number of distinct page breaks may cause high memory consumption**.


### Styling

Styling lets you apply colors or other formatting to cells being written. A style is made up of five components:
- the **font**, including face, size and text color, represented by an `XlsxFont` object
- the **fill**, specifying the background color, represented by an `XlsxFill` object
- the **border** style and color, represented by an `XlsxBorder` object
- the **number format**, specifying how a number should appear, such as how many decimals or whether to show it as percentage or date and time, represented by an `XlsxNumberFormat` object
- the **alignment**, specifying horizontal and vertical alignment of cell content or rotation, represented by an `XlsxAlignment` object.

You can create a new style, combining the above five elements, using the constructor of the `XlsxStyle` class, or by cloning an existing style replacing one element with a `With` method:

```csharp
// class XlsxStyle
public XlsxStyle(XlsxFont font, XlsxFill fill, XlsxBorder border, XlsxNumberFormat numberFormat,
                 XlsxAlignment alignment);
public XlsxStyle With(XlsxFont font);
public XlsxStyle With(XlsxFill fill);
public XlsxStyle With(XlsxBorder border);
public XlsxStyle With(XlsxNumberFormat numberFormat);
public XlsxStyle With(XlsxAlignment alignment);
```

The resulting `XlsxStyle` object can be used with a `Write` method to style a cell being written, or with `SetDefaultStyle` to change the default style of the `XlsxWriter`, or to specify column or row styles.

Under the hood, all styles used with an `XlsxWriter` are collected in a stylesheet, and this library **deduplicates** them in **constant time** as needed. Thus, you should usually not worry about performance or memory consumption when you use multiple styles. Note, however, that the stylesheet is kept in RAM until the `XlsxWriter` is disposed, thus using a *very* large number of *different* styles may cause high memory consumption.

The built-in `XlsxStyle.Default` object provides a ready-to-use style combining the built-in `XlsxFont.Default` font, the built-in `XlsxFill.None` fill, the built-in `XlsxBorder.None` border, the built-in `XlsxNumberFormat.General` number format and the built-in `XlsxAlignment.Default` alignment. All those elements are explained later.

#### The default style

Each `XlsxWriter` has a default style that is used whenever you write a cell using `Write` without specifying an explicit style. To read the current default style, or set it while fluently writing the XLSX file, use:

```csharp
// class XlsxWriter
public XlsxStyle DefaultStyle { get; }
public XlsxWriter SetDefaultStyle(XlsxStyle style);
```

The default style of a new `XlsxWriter` is set to `XlsxStyle.Default`, but you can change it at any time, including setting it back to `XlsxStyle.Default`.


#### Fonts

An `XlsxFont` object lets you define the font face, its size in points, the text color and emphasis such as bold, italic and strikeout. Create it via the constructor, or clone an existing one replacing a property with a `With` method:

```csharp
// class XlsxFont
public XlsxFont(
        string name,
        double size,
        System.Drawing.Color color,
        bool bold = false,
        bool italic = false,
        bool strike = false,
        XlsxFont.Underline underline = XlsxFont.Underline.None);
public XlsxFont With(System.Drawing.Color color);
public XlsxFont WithName(string name);
public XlsxFont WithSize(double size);
public XlsxFont WithBold(bool bold = true);
public XlsxFont WithItalic(bool italic = true);
public XlsxFont WithStrike(bool strike = true);
public XlsxFont WithUnderline(XlsxFont.Underline underline = XlsxFont.Underline.None);
```

XlsxFont.Underline enum provides underline styles None, Single, Double, SingleAccounting and DoubleAccounting.

Using named arguments is recommended to improve readability. For example to create a red, italic, 11-point, Calibri font, use:\
`var redItalicFont = new XlsxFont("Calibri", 11, Color.Red, italic: true)`.

The built-in `XlsxFont.Default` object provides a default black, plain, 11-point, Calibri font.

#### Fills

An `XlsxFill` object lets you define the background color of a cell and the pattern to use to fill the background. Create it with:

```csharp
// class XlsxFill
public XlsxFill(System.Drawing.Color color, XlsxFill.Pattern patternType = XlsxFill.Pattern.Solid);
```

The `XlsxFill.Pattern` enum defines how to apply the background color, and may be `None` for a transparent fill, `Solid` for a solid fill or `Gray125` for a dotted pattern (not necessarily gray) with 12.5% coverage. For example to create a yellow solid fill, use:\
`var yellowFill = new XlsxFill(Color.Yellow)`

The built-in `XlsxFill.None` object provides a default empty fill.

#### Borders

An `XlsxBorder` object lets you define thickness and colors for the borders of a cell. Create a new border set via one of the following:

```csharp
// class XlsxBorder
public XlsxBorder(
        XlsxBorder.Line top = null,
        XlsxBorder.Line right = null,
        XlsxBorder.Line bottom = null,
        XlsxBorder.Line left = null,
        XlsxBorder.Line diagonal = null,
        bool diagonalDown = false,
        bool diagonalUp = false);
public static XlsxBorder Around(XlsxBorder.Line around);
```

The constructor lets you specify each border individually. Using named arguments is recommended to improve readability. The two diagonal borders share the same line style, but you can choose whether to show them individually. The `Around` named constructor is a shortcut for the common case of setting the top, right, bottom and left borders to the same line style, with no diagonals.

Each line is constructed with:

```csharp
// class XlsxBorder.Line
public Line(System.Drawing.Color color, XlsxBorder.Style style);
```

The `XlsxBorder.Style` enum defines the stroke style, such as `None`, `Thin`, `Thick`, `Dashed` and others.

For example to create a thin black border on the left side only, use:\
`var leftBorder = new XlsxBorder(left: new XlsxBorder.Line(Color.Black, XlsxBorder.Style.Thin))`.

The built-in `XlsxBorder.None` object provides a default empty border set.

#### Number formats

An `XlsxNumberFormat` object lets you define how the content of a cell should appear when it contains a numeric value. Create it via the constructor:

```csharp
// class XlsxNumberFormat
public XlsxNumberFormat(string formatCode);
```

The format code has the same format you would normally use in Excel, such as `"0.0%"` for a percentage with exactly one decimal value. For example, to create a custom number format with a thousand separator, at least two decimal digits and at most six, use:\
`var customNumberFormat = new XlsxNumberFormat("#,##0.00####")`

Excel defines and reserves many "magic" number formats, and this library exposes some of them as:

* `XlsxNumberFormat.General`: the default number format, where Excel automatically chooses the "best" representation based on magnitude and number of decimals
* `XlsxNumberFormat.Integer`: no decimal digits, that is the format code `"0"`
* `XlsxNumberFormat.TwoDecimal`: two decimal digits, that is the format code `"0.00"`
* `XlsxNumberFormat.ThousandInteger`: thousand separators and no decimal digits, that is the format code `"#,##0"`
* `XlsxNumberFormat.ThousandTwoDecimal`: thousand separators and two decimal digits, that is the format code `"#,##0.00"`
* `XlsxNumberFormat.IntegerPercentage`: percentage formatting and no decimal digits, that is the format code `"0%"`
* `XlsxNumberFormat.TwoDecimalPercentage`: percentage formatting and two decimal digits, that is the format code `"0.00%"`
* `XlsxNumberFormat.Scientific`: scientific notation with two decimals and two-digit exponent, that is the format code `"0.00E+00"`
* `XlsxNumberFormat.ShortDate`: localized day, month and year as digits; for a European format the equivalent code would be `"dd/mm/yyyy"` but the actual code would be locale-dependent
* `XlsxNumberFormat.ShortDateTime`: localized day, month and year as digits with hours and minutes; for a European format the equivalent code would be `"dd/mm/yyyy hh:mm"` but the actual code would be locale-dependent
* `XlsxNumberFormat.Text`: treat newly inserted numbers as text, that is the format code `"@"`


#### Alignment

An `XlsxAlignment` object describes alignment and other text control properties, constructed with:

```csharp
public XlsxAlignment(
        XlsxAlignment.Horizontal horizontal = XlsxAlignment.Horizontal.General,
        XlsxAlignment.Vertical vertical = XlsxAlignment.Vertical.Bottom,
        int indent = 0,
        bool justifyLastLine = false,
        XlsxAlignment.ReadingOrder readingOrder = XlsxAlignment.ReadingOrder.ContextDependent,
        bool shrinkToFit = false,
        int textRotation = 0,
        bool wrapText = false);
```

Using named arguments is recommended to improve readability. The parameters represent:
- `horizontal`: horizontal alignment of the text, such as left, right, center or justified
- `vertical`: vertical alignment of the text, such as top, bottom, center or justified
- `indent`: how many spaces the cell content must be indented
- `justifyLastLine`: whether to justify even the last line when the alignment is set to `Justify`
- `readingOrder`: the text direction such as left-to-right or right-to-left
- `shrinkToFit`: whether to reduce automatically the font size to fit the content into the cell
- `textRotation`: rotation angle in degrees of the cell content, in range 0..180
- `wrapText`: whether to insert line breaks automatically into the text to fit the content into the cell

The built-in `XlsxAlignment.Default` object provides an alignment with default values for all properties.


## Special thanks

Kudos to [Roberto Montinaro](https://github.com/montinaro), [Matteo Pierangeli](https://github.com/matpierangeli) and [Giovanni Improta](https://github.com/improtag) for their patience and their very valuable suggestions! <3

Many thanks to all the great people who contributed with [pull requests](https://github.com/salvois/LargeXlsx/pulls?q=), questions and reports!


## License

Permissive, [2-clause BSD style](https://opensource.org/licenses/BSD-2-Clause)

LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020-2025  Salvatore ISAJA

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
