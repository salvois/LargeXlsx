# LargeXlsx - Minimalistic .net library to write large XLSX files

This is a minimalistic library, written in C# targeting .net standard 2.0, providing a tiny layer above Microsoft's [Office Open XML library](https://github.com/OfficeDev/Open-XML-SDK) to facilitate creation of very large Excel files in XLSX format.

This library provides simple primitives to write data in a streamed manner, so that potentially huge files can be created while consuming a low, constant amount of memory.


## Supported features

Currently the library supports:

* cells containing inline strings and numeric (double) values
* multiple worksheets
* merged cells
* split panes, a.k.a. frozen rows and columns
* basic styling: font face, size, and color; solid background color; top-right-bottom-left cell borders; numeric formatting


## Example

To create a simple single-sheet Excel document:

```csharp
using (var stream = new FileStream("Simple.xlsx", FileMode.Create))
using (var xlsxWriter = new XlsxWriter(stream))
{
    var whiteFont = xlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ffffff", bold: true);
    var blueFill = xlsxWriter.Stylesheet.CreateSolidFill("004586");
    var headerStyle = xlsxWriter.Stylesheet.CreateStyle(
            whiteFont, blueFilll, XlsxBorder.None, XlsxNumberFormat.General);

    xlsxWriter.BeginWorksheet("Sheet1")
        .BeginRow().Write("Col1", headerStyle).Write("Col2", headerStyle).Write("Col3", headerStyle)
        .BeginRow().Write("Row2").Write(42).Write(-1)
        .BeginRow().Write("Row3").SkipColumns(1).Write(1234)
        .SkipRows(2)
        .BeginRow().AddMergedCell(1, 2).Write("Row6").SkipColumns(1).Write(3.14159265359);
}
```

The output is like:

![Single sheet Excel document with 6 rows and 3 columns](https://github.com/salvois/LargeXlsx/raw/master/example.png)

## Known issues

On .net core there is an [issue on System.IO.Packaging](https://github.com/dotnet/corefx/issues/24457) (used by the Open XML SDK to write XLSX's zip packages) that causes memory consumption to be proportional to the amount of data written, instead of being low and constant. Unfortunately, this kind of defeats the purpose of this library when targeting .net core. The issue is not present on .net framework.

## Usage

The `XlsxWriter` class is the entry point for all functionality of the library. It is designed so that most of its methods can be chained to write the Excel file using a fluent syntax.
Please note that an `XlsxWriter` object **must be disposed** to properly finalize the Excel file. Sandwitching its lifetime in a `using` statement is recommended.
Pass the constructor a `Stream` to save the Excel file into. Please note that, due to internals of the Office Open XML library, the stream must be opened for both read and write.

    XlsxWriter(Stream stream)

The recipe is adding a worksheet with `BeginWorksheet`, adding a row with `BeginRow`, writing cells to that row with `Write`, and repeating as required. Rows and worksheets are implicitly finalized as soon as new rows or worksheets are added, or the `XlsxWriter` is disposed.

### The insertion point

To enable streamed write, the content of the Excel file must be written strictly from top to bottom and from left to right. Think of an insertion point always advancing when writing content.
The `CurrentRowNumber` and `CurrentColumnNumber` read-only properties will return the location of the **next** cell that will be written. Both the row and column numbers are **one-based**.

    int CurrentRowNumber { get; }
    int CurrentColumnNumber { get; }

Please note that `CurrentColumnNumber` may be zero, and thus invalid, if the current row has not been set up using `BeginRow` (attempting to write a cell would throw an exception).

### Creating a new worksheet

Call `BeginWorksheet` passing the sheet name and, optionally, the one-based indexes of the row and column where to place a split to create frozen panes.
A call to `BeginWorksheet` finalizes the last worksheet being written, if any, and sets up a new one, so that rows can be added.

    XlsxWriter BeginWorksheet(string name, int splitRow = 0, int splitColumn = 0)

### Adding or skipping rows

Call `BeginRow` to advance the insertion point to the beginning of the next line and set up a new row to accept content. If a previous row was being written, it is finalized before creating the new one.

    XlsxWriter BeginRow()

Call `SkipRows` to move the insertion point down by the specified count of rows, that will be left empty and unstylized. If a previous row was being written, it is finalized. Please note that `BeginRow` must be called anyways before starting to write a new row.

    XlsxWriter SkipRows(int rowCount)

### Writing cells

Call one of the `Write` overloads to write content to the cell at the insertion point. You may write one of the following:

  * **Nothing**: a cell containing no value, that will usually be deserialized as `null`.
  * **Inline string**: a string of text that is written directly into the cell; this is in contrast with a different functionality of the XLSX file format, which can support a global look-up table of strings, and just the string index into the cell; the latter functionality is not supported because it is inherently incompatible with streamed write. If the string is `null` the method falls back on the "Nothing" case.
  * **Number**: a numeric constant, that will be interpreted as a `double` value; conveniency overloads accepting `int` and `decimal` are provided, but the under the hood the value will be converted to `double` because it is the only numeric type truly supported by the XLSX file format.


    XlsxWriter Write()
    XlsxWriter Write(string value)
    XlsxWriter Write(double value)
    XlsxWriter Write(decimal value)
    XlsxWriter Write(int value)

Besides the value to write into the cell, `Write` optionally accepts another parameter representing the ID of the style (see Styling) to use to stylize the cell being written. This cannot be changed after the cell has been written.

    XlsxWriter Write(XlsxStyle style)
    XlsxWriter Write(string value, XlsxStyle style)
    XlsxWriter Write(double value, XlsxStyle style)
    XlsxWriter Write(decimal value, XlsxStyle style)
    XlsxWriter Write(int value, XlsxStyle style)

Like rows, cells can be skipped using the `SkipColumns` method, to move the insertion point to the right by the specified count of cells, that will be left empty and unstylized.

    XlsxWriter SkipColumns(int columnCount)

### Merged cells

A rectangle of adjacent cells can be merged using the `AddMergedCells` method. Content for the merged cells must be written in the top-left cell of the rectangle, and the other cells of the merged rectangle must be explicitly skipped using `SkipColumns` and/or `SkipRows` as appropriate.

    XlsxWriter AddMergedCell(int fromRow, int fromColumn, int rowCount, int columnCount)

For example, if merging the 2 rows x 3 columns range `A7:C8` using `AddMergedCells(7, 1, 2, 3)`, you must write content for the merged cell in `A7`, then explicitly jump by further 2 columns using `SkipColumns(2)` to continue writing content from `D7`, and the same applies on row 8, where after a `BeginRow()` you must skip 3 columns with `SkipColumns(3)` and continue writing from `D8`.

**Note**: due to the structure of the XLSX file format, the ranges for all merged cells of a worksheet must be accumulated in RAM, because they must be written to the file after the content of the whole worksheet. **Using a large number of merged cells may cause high memory consumption**.
This also means that you may call `AddMergedCells` at any moment while you are writing a worksheet (that is between a `BeginWorksheet` and the next one, or disposal of the `XlsxWriter` object), even for cells already written or well before writing them.

To facilitate merging cells while fluently writing the file, a conveniency overload is provided, using the insertion point as the top-left cell for the merged range, thus requiring you to only specify the height and width of the merged rectangle. This does not advance the insertion point, thus a `Write` should usually follow to write content for the merged rectangle, followed by `SkipColumns` as needed (see the last row of the Example above).

    XlsxWriter AddMergedCell(int rowCount, int columnCount)

### Styling

Styling lets you apply colors or other formatting to cells being written. The XLSX file format uses the concept of **stylesheet** where you list all possible styles used by your content, and each style is identified by a **style ID**, that is an index in the table of styles that makes up the stylesheet, represented as an `XlsxStyle` value.
When you write a cell using `Write` you can specify the style ID to use for that cell.

Each `XlsxWriter` object manages a single `XlsxStylesheet` object, exposed by the `Stylesheet` property, that provides functionality to add styles to the stylesheet.

    XlsxStylesheet Stylesheet { get; }

A style is made up of four components: the **font**, including face, size and text color; the **fill**, specifying the background color; the **border** style and color; the **number format** specifying how a number should appear, such as how many decimals or whether to show it as percentage.
The stylesheet contains a table for each of the four style component types, thus you will also have **font IDs**, **fill IDs**, **border IDs** and **number format IDs**, represented as `XlsxFont`, `XlsxFill`, `XlsxBorder` and `XlsxNumberFormat` values respectively.

**Note:** due to the internals of the XLSX file format, the content of the stylesheet (thus the list of fonts, fills, borders, number formats and the styles combining them) are kept in RAM until the `XlsxWriter` is disposed. **Using a large number of styles may cause high memory consumption**.

#### Fonts

To create a new font, call the `CreateFont` method of the `XlsxStylesheet` object. Using named arguments is recommended to improve readability. The color is a string of hexadecimal digits in RRGGBB format.

    XlsxFont CreateFont(string fontName, double fontSize, string hexRgbColor,
                        bool bold = false, bool italic = false, bool strike = false)

For example to create a red, italic, 11-point Calibri font, use: `var redItalicFont = xlsxWriter.Stylesheet.CreateFont("Calibri", 11, "ff0000", italic: true)`.

A default black, plain, 11-point Calibri font, can be referenced with font ID is `XlsxFont.Default`.

### Fills

Currently, only fills with a solid background color are supported. To create a new solid fill, call the `CreateSolidFill` method of the `XlsxStylesheet` object. The color is a string of hexadecimal digits in RRGGBB format.

    XlsxFill CreateSolidFill(string hexRgbColor)
    
For example to create a yellow fill, use: `var yellowFill = xlsxWriter.Stylesheet.CreateSolidFill("ffff00")`.

A default empty fill can be referenced with the fill ID `XlsxFill.None`.

### Borders

Currently, a set of top, right, bottom and left cell borders of the same color is supported. To create a new set of borders, call the `CreateBorder` method of the `XlsxStylesheet` object. The color is a string of hexadecimal digits in RRGGBB format. The `BorderStyleValues` enum from the Office Open XML library defines the kind of border of each cell side, such as `None`, `Thin`, `Medium`, `Thick`, `Double`, `Dashed`, `Dotted` and others. Using named arguments is recommended to improve readability.

    public XlsxBorder CreateBorder(
            string hexRgbColor,
            BorderStyleValues top = BorderStyleValues.None,
            BorderStyleValues right = BorderStyleValues.None,
            BorderStyleValues bottom = BorderStyleValues.None,
            BorderStyleValues left = BorderStyleValues.None)
                                 
For example to create a thin black border on the left side only, use: `var leftBorder = xlsxWriter.Stylesheet.CreateBorder("000000", left: BorderStyleValues.Thin)`.

A default empty border set can be referenced with the border ID `XlsxBorder.None`.

### Number formats

To create a custom number format, call the `CreateNumberFormat` method of the `XlsxStylesheet` object, specifying a number format string as you would normally do in Excel, such as `"0.0%"` for a percentage with exactly one decimal value.

    XlsxNumberFormat CreateNumberFormat(string formatCode)
    
For example to create a custom number format with thousand separator, at least two decimal digits and at most six, use: `var customNumberFormat = xlsxWriter.Stylesheet.CreateNumberFormat("#,##0.00####")`.

Excel defines and reserves many number formats, and this library exposes some of their number format IDs as:

* `XlsxNumberFormat.General`: the default number format, where Excel automatically chooses the "best" representation based on magnitude and number of decimals.
* `XlsxNumberFormat.TwoDecimal`: a number format with thousand separators and two decimal numbers, that is the format code `"#,##0.00"`.

### Combining them all to create a style

To create a new style using the specified combination of font, fill, border and number format, call the `CreateStyle` method of the `XlsxStylesheet` object. The resulting style ID can be used with `Write` to stylize a cell being written.

    XlsxStyle CreateStyle(XlsxFont font,
                          XlsxFill fill,
                          XlsxBorder border,
                          XlsxNumberFormat numberFormat)

This library provides the style ID `XlsxStyle.Default` for a deafult style combining `XlsxFont.Default`, `XlsxFill.None`, `XlsxBorder.None` and `XlsxNumberFormat.General`. This is the style ID used whenever `Write` is called without an explicit style ID parameter.

## Special thanks

Kudos to [Roberto Montinaro](https://github.com/montinaro) and [Matteo Pierangeli](https://github.com/matpierangeli) for their patience and their very valuable suggestions! <3

## License

Permissive, [2-clause BSD style](https://opensource.org/licenses/BSD-2-Clause)

LargeXlsx - Minimalistic .net library to write large XLSX files

Copyright 2020  Salvatore ISAJA

Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.

2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
