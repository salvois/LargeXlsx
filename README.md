# LargeXlsx - A .net library to write large XLSX files

[![NuGet](https://img.shields.io/nuget/v/LargeXlsx.svg)](https://www.nuget.org/packages/LargeXlsx)

This is a minimalistic yet feature-rich library, written in C# targeting .net standard 2.0, providing simple primitives to write Excel files in XLSX format in a streamed manner, so that potentially huge files can be created while consuming a low, constant amount of memory.


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


## Changelog

* 1.7: Write overload for booleans, more performance improvements thanks to [Antony Corbett](https://github.com/AntonyCorbett)
* 1.6: Opt-in shared string table (for memory vs. file size trade-off)
* 1.5: Password protection of sheets
* 1.4: Support for font underline, thanks to [Sergey Bochkin](https://github.com/sbochkin)
* 1.3: Optional ZIP64 support for really huge files
* 1.2: Right-to-left worksheets. Parametric compression level to trade between space and speed
* 1.1: Number format for text-formatted cells (the "@" formatting)
* 1.0: Finalized API
* 0.0.9: Started writing XLSX files directly rather than using Microsoft's [Office Open XML library](https://github.com/OfficeDev/Open-XML-SDK)

## Usage

The `XlsxWriter` class is the entry point for almost all functionality of the library. It is designed so that most of its methods can be chained to write the Excel file using a fluent syntax.

The constructor allows you to create an XLSX writer. Please note that an `XlsxWriter` object **must be disposed** to properly finalize the Excel file. Sandwitching its lifetime in a `using` statement is recommended.

```csharp
// class XlsxWriter
public XlsxWriter(
    Stream stream,
    SharpCompress.Compressors.Deflate.CompressionLevel compressionLevel = CompressionLevel.Level3, // compressionLevel since version 1.2
    bool uzeZip64 = false); // useZip64 since version 1.3
```

The constructor accepts:
* A writeable `Stream` to save the Excel file into
* An optional desired compression level of the underlying zip stream. The default `CompressionLevel.Level3` roughly matches file sizes produced by Excel. Higher compression levels may result in lower speed.
* An optional flag indicating whether to use ZIP64 compression to support content larger than 4 GiB uncompressed. Recent versions of XLSX-enabled applications such as Excel or LibreOffice should be able to read any file compressd using ZIP64, even small ones, thus, if you don't know the file size in advance and you target recent software, you could just set it to `true`.

The recipe is adding a worksheet with `BeginWorksheet`, adding a row with `BeginRow`, writing cells to that row with `Write`, and repeating as required. Rows and worksheets are implicitly finalized as soon as new rows or worksheets are added, or the `XlsxWriter` is disposed.

### The insertion point

To enable streamed write, the content of the Excel file must be written strictly from top to bottom and from left to right. Think of an insertion point always advancing when writing content.\
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


### Creating a new worksheet

Call `BeginWorksheet` passing the sheet name and, optionally, the one-based indexes of the row and column where to place a split to create frozen panes. Setting `rightToLeft` to `true` switches the worksheet to right-to-left mode (to support languages such as Arabic and Hebrew). Finally, the `columns` parameter can be used to specify optional column formatting.\
A call to `BeginWorksheet` finalizes the last worksheet being written, if any, and sets up a new one, so that rows can be added.

```csharp
// class XlsxWriter
public XlsxWriter BeginWorksheet(
        string name,
        int splitRow = 0,
        int splitColumn = 0,
        bool rightToLeft = false, // rightToLeft since version 1.2
        IEnumerable<XlsxColumn> columns = null);
```

Note that, for compatibility with a restriction of the Excel application, names are restricted to a maximum of 31 character. An `ArgumentException` is thrown if a longer name is passed.
An `ArgumentException` is also thrown when trying to add a worksheet with a name already used for another worksheet.

#### Column formating

The `BeginWorksheet` method accepts an optional `columns` parameter to specify a list of column formatting objects of type `XlsxColumn`, each describing one or more adjacent columns, with their custom width, hidden state or default style, starting from column A.

This information must be provided before writing any content to the worksheet, thus the number, width and styles of the columns must be known in advance.

You can create `XlsxColumn` objects with one of these named constructors:

```csharp
// class XlsxColumn
public static XlsxColumn Unformatted(int count = 1);
public static XlsxColumn Formatted(double width, int count = 1, bool hidden = false, XlsxStyle style = null);
```

`Unformatted` creates a column description that is used basically to skip one or more unformatted columns.

`Formatted` creates a column description to specify the mandatory witdh, optional hidden state, and optional style of one or more contiguous columns. The width is expressed (simplyfing) in approximate number of characters. The column style represents how to style all *empty* cells of a column. Cells that are explicitly written always use the cell style instead.

### Adding or skipping rows

Call `BeginRow` to advance the insertion point to the beginning of the next line and set up a new row to accept content. If a previous row was being written, it is finalized before creating the new one.

```csharp
// class XlsxWriter
public XlsxWriter BeginRow(double? height = null, bool hidden = false, XlsxStyle style = null);
```
    
You can specify optional row formatting when creating a new row. The height is expressed in points. The row style represent how to style all *empty* cells of a row. Cells that are explicitly written always use the cell style instead.

Call `SkipRows` to move the insertion point down by the specified count of rows, that will be left empty and unstyled (unless column styles are in place). If a previous row was being written, it is finalized. Please note that `BeginRow` must be called anyways before starting to write a new row.

```csharp
// class XlsxWriter
public XlsxWriter SkipRows(int rowCount);
```

### Writing cells

Call one of the `Write` methods to write content to the cell at the insertion point:

```csharp
// class XlsxWriter
public XlsxWriter Write(XlsxStyle style = null, int columnSpan = 1, int repeatCount = 1);
public XlsxWriter Write(string value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(double value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(decimal value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(int value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(DateTime value, XlsxStyle style = null, int columnSpan = 1);
public XlsxWriter Write(bool value, XlsxStyle style = null, int columnSpan = 1); // since version 1.7
public XlsxWriter WriteFormula(string formula, XlsxStyle style = null, int columnSpan = 1, IConvertible result = null);
public XlsxWriter WriteSharedString(string value, XlsxStyle style = null, int columnSpan = 1); // since version 1.6
```

 You may write one of the following:

  * **Nothing**: a cell containing no value, but styled nonetheless.
  * **String**: an inline literal string of text; if the string is `null` the method falls back on the "nothing" case. The value of an inline string is written into the cell, thus resulting in low memory consumption but possibly larger files (see, in contrast, shared strings).
  * **Number**: a numeric constant, that will be interpreted as a `double` value; convenience overloads accepting `int` and `decimal` are provided, but under the hood the value will be converted to `double` because it is the only numeric type truly supported by the XLSX file format.
  * **Date and time**: a `DateTime` value, that will be converted to its `double` representation (days since 1900-01-01). Note that you must style the cell using a date/time number format to have the value appear as a date.
  * **Boolean**: a `bool` value, that will appear either as `TRUE` or `FALSE`
  * **Formula**: a string that Excel or a compatible application will interpret as a formula to calculate. Note that, unless you provide a `result` calculated by yourself (either string or numeric), no result is saved into the XLSX file. However, a spreadsheet application will calculate the result as soon as the XLSX file is opened.
  * **Shared string**: a shared literal string of text. Contrary to inline strings, shared strings are saved in a look-up table and deduplicated in constant time, and only a reference is written into the cell. This may help to produce smaller files, but, since the shared strings must be accumulated in RAM, **writing a large number of *different* shared strings may cause high memory consumption**. Keep this in mind when choosing between inline strings and shared strings.

The `style` parameter specifies the style to use for the cell being written. If `null` (or omitted), the cell is styled using the current default style of the `XlsxWriter` (see Styling). Note that in no case the style of the column or the row, if any, is used for written cells.

When writing empty cells you can also specify a `repeatCount` parameter, to write multiple consecutive styled empty cells. Note that the difference between `repeatCount` and `columnSpan` both greater than 1 is that the latter creates a merged cell (with its memory consumption drawback) and does not have in-between borders.

The `columnSpan` parameter can be used to let the cell span multiple columns. When greater than 1, a merged range of such cells is created (see Merged cells), content is written to the first cell of the range and the insertion point is advanced after the merged cells. Note that `xlsxWriter.Write(value, columnSpan: count)` is actually a shortcut for `xlsxWriter.AddMergedCells(1, count).Write(value).Write(repeatCount: count - 1)`. Since a merged cell is created, **writing a large number of cells with columnSpan greater than 1 may cause high memory consumption**.

#### Writing hyperlinks

The XLSX file format provides two ways to insert hyperlinks (either to a web site or cells perhaps in a differet workbook): using a specific "hyperlinks" section in the worksheet and using the `HYPERLINK` formula in a cell.

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
- `prompt`: an optional message to be shown as tooltip when the cell receives focus
- `promptTitle`: an optional title to show in the prompt tooltip
- `showDropDown`: whether to show a dropdown list when the validation type is set to `XlsxDataValidation.ValidationType.List`; **Note:** for some reason, this seems to work backwards, with both Excel and LibreOffice showing the dropdown when this property is **false**
- `showErrorMessage`: whether the spreadsheet application will show an error when invalid content is detected
- `showInputMessage`: whether the spreadsheet application will show the prompt tooltip when a validated cell is focused
- `validationType`: specifies to do validation on numeric values, whole integer values, date values, text length (using the comparison operator specified by `operatorType`) or against a list of allowed values
- `formula1`: the value to compare the cell content against; for lists, it can be a reference to a cell range containing the allowed values, or a list of comma separated case-sensitive string contants, enclosed by double quotes (e.g. `"\"item1,item2\"")`; for other validation types, it can be a reference to a cell containing the value or a constant
- `formula2`: for comparison involving two values (e.g. operator `Between`), the second value, specified as in `formula1`.

To ease creation of an `XlsxDataValidation` specifying validation against a list of string contants, the following named constructor is provided. Note that it basically does a `string.Join` of choices into `formula1`, thus if a choice includes a comma, the spreadsheet application will split it in separate choices. There is no way to include real commas in choices specified as string constants (rather than as cell range).

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

Password protection of worksheets helps preventing accidental modification of data. You can enable protection on the worksheet being written using:

```csharp
// class XlsxWriter
public XlsxWriter SetSheetProtection(XlsxSheetProtection sheetProtection); // since version 1.5

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

**Note:** password protection of sheets is not to be confused with workbook encryption and is not meant to be secure. File contents are still written in clear text and may be changed by deliberately editing the file. The password is not written into the file but a hash of the password is.


### Styling

Styling lets you apply colors or other formatting to cells being written. A style is made up of five components:
- the **font**, including face, size and text color, represented by an `XlsxFont` object
- the **fill**, specifying the background color, represented by an `XlsxFill` object
- the **border** style and color, represented by an `XlsxBorder` object
- the **number format** specifying how a number should appear, such as how many decimals or whether to show it as percentage or date and time, represented by an `XlsxNumberFormat` object
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

Each `XlsxWriter` has a default style that is used whenever you write a cell using `Write` without specifying an explicit style. To read the current default style, or set it while fluently write the XLSX file, use:

```csharp
// class XlsxWriter
public XlsxStyle DefaultStyle { get; }
public XlsxWriter SetDefaultStyle(XlsxStyle style);
```

The default style of a new `XlsxWriter` is set to `XlsxStyle.Default`, but you can change it at any time, including setting it back to `XlsxStyle.Default`.


#### Fonts

An `XlsxFont` object lets you define the font face, its size in points, the text color and emphasis such as bold, italic and strikeout. Create it via constructor, or clone an existing one replacing a property with a `With` method:

```csharp
// class XlsxFont
public XlsxFont(
        string name,
        double size,
        System.Drawing.Color color,
        bool bold = false,
        bool italic = false,
        bool strike = false,
        XlsxFont.Underline underline = XlsxFont.Underline.None); // underline since version 1.4
public XlsxFont With(System.Drawing.Color color);
public XlsxFont WithName(string name);
public XlsxFont WithSize(double size);
public XlsxFont WithBold(bool bold = true);
public XlsxFont WithItalic(bool italic = true);
public XlsxFont WithStrike(bool strike = true);
public XlsxFont WithUnderline(XlsxFont.Underline underline = XlsxFont.Underline.None); // since version 1.4
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

An `XlsxNumberFormat` object lets you define how the content of a cell should appear when it contains a numeric value. Create it via constructor:

```csharp
// class XlsxNumberFormat
public XlsxNumberFormat(string formatCode);
```

The format code has the same format you would normally use in Excel, such as `"0.0%"` for a percentage with exactly one decimal value. For example, to create a custom number format with thousand separator, at least two decimal digits and at most six, use:\
`var customNumberFormat = new XlsxNumberFormat("#,##0.00####")`.

Excel defines and reserves many "magic" number formats, and this library exposes some of them as:

* `XlsxNumberFormat.General`: the default number format, where Excel automatically chooses the "best" representation based on magnitude and number of decimals.
* `XlsxNumberFormat.Integer`: no decimal digits, that is the format code `"0"`.
* `XlsxNumberFormat.TwoDecimal`: two decimal digits, that is the format code `"0.00"`.
* `XlsxNumberFormat.ThousandInteger`: thousand separators and no decimal digits, that is the format code `"#,##0"`.
* `XlsxNumberFormat.ThousandTwoDecimal`: thousand separators and two decimal digits, that is the format code `"#,##0.00"`.
* `XlsxNumberFormat.IntegerPercentage`: percentage formatting and no decimal digits, that is the format code `"0%"`.
* `XlsxNumberFormat.TwoDecimalPercentage`: percentage formatting and two decimal digits, that is the format code `"0.00%"`.
* `XlsxNumberFormat.Scientific`: scientific notation with two decimals and two-digit exponent, that is the format code `"0.00E+00"`.
* `XlsxNumberFormat.ShortDate`: localized day, month and year as digits; for a European format the equivalent code would be `"dd/mm/yyyy"` but the actual code would be locale-dependent.
* `XlsxNumberFormat.ShortDateTime`: localized day, month and year as digits with hours and minutes; for a European format the equivalent code would be `"dd/mm/yyyy hh:mm"` but the actual code would be locale-dependent.
* `XlsxNumberFormat.Text`: treat newly inserted numbers as text, that is the format code `"@"` (since version 1.1).


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

## License

Permissive, [2-clause BSD style](https://opensource.org/licenses/BSD-2-Clause)

LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020  Salvatore ISAJA

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
