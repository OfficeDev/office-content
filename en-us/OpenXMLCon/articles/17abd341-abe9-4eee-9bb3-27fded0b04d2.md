<!--This is the start of the document-->
# Working with formulas (Open XML SDK)
**Last modified:** April 11, 2013

_**Applies to:** Office 2013 | Open XML_

This topic discusses the Open XML SDK 2.5  **CellFormula** class and how it relates to the Open XML File Format SpreadsheetML schema. For more information about the overall structure of the parts and elements that make up a SpreadsheetML document, see [Structure of a SpreadsheetML document (Open XML SDK)](3b35a153-c8ff-4dc7-96d5-02c515f31770.md).


## Formulas in SpreadsheetML
You can use formulas to create computational models. Formulas allow for automatic calculation of values based on data inside and outside the spreadsheet or the output of other computed cells in the spreadsheet.

Formulas are stored inside each cell that uses a formula, in the worksheet XML file. Use the  **CellFormula** (< **f**>) element to define the formula text. Formulas can contain mathematical expressions that include a wide range of predefined functions.

The  **CellValue** (< **v**>) element, stores the cached formula value based on the last time the formula was calculated. This allows the user to postpone calculation of the formula values when the spreadsheet is opened, which saves time when opening a worksheet. You do not have to specify the value, and if you omit it, it is the responsibility of the Open XML reader to compute the value based on the formula definition when the worksheet is opened. For more information about the  **CellValue** class, see **CellValue**.

The following information from the  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification introduces the **cellFormula** (< **f**>) element.

A SpreadsheetML formula is the syntactic representation of a series of calculations that is parsed or interpreted by the spreadsheet application into a function that calculates a value or array of values based upon zero-to-many inputs.

A formula is an expression that can contain the following: constants, operators, cell references, calls to functions, and names.

[Example: Consider the formula PI()*(A2^2). In this case,

 PI() results in a call to the function PI, which returns the value of p.

 The cell reference A2 returns the value in that cell.

 2 is a numeric constant.

 The caret (^) operator raises its left operand to the power of its right operand.

 The parentheses, ( and ), are used for grouping.

 The asterisk (*) operator performs multiplication of its two operands.

end example]

An operator is a symbol that specifies the type of operation to perform on one or more operands. There are arithmetic, comparison, text, and reference operators.

Each set of horizontal cells in a worksheet is a row, and each set of vertical cells is a column. A cell's row and column combination designates the location of that cell.

A cell reference designates one or more cells on the same worksheet. Using references, one can:

 Use data contained in different parts of the same worksheet in a single formula.

 Use the value from a single cell in several formulas.

 Refer to cells on other sheets in the same workbook, and even to other workbooks. (References to cells in other workbooks are called links.)

A function is a named formula that takes zero or more arguments, performs an operation, and, optionally, returns a result. Some examples of function calls are: PI(), POWER(A1,B3), and SUM(C6:C10).

There are more than 300 predefined functions defined by this Office Open XML specification. User-defined functions are also permitted.

A name is an alias for a constant, a cell reference, or a formula. A name in a formula can make it easier to understand the purpose of that formula. For example, the formula SUM(FirstQuarterSales) is easier to identify than SUM(C20:C30).

Each expression has a type. SpreadsheetML formulas support the following types: array, error, logical, number, and text.

An array value or constant represents a collection of one or more elements, whose values can have any type (i.e., the elements of an array need not all have the same type).

© ISO/IEC29500: 2008.

For more information about formula syntax see the ISO/IEC 29500 specification.


### SpreadsheetML Example
This example shows the XML for a file that contains one formula, the SUM function, in cell A6 on Sheet1. The following XML defines the worksheet and is contained in the "sheet1.xml" file.


```XML
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">
    <dimension ref="A1:A6"/>
    <sheetViews>
        <sheetView tabSelected="1" workbookViewId="0">
            <selection activeCell="A7" sqref="A7"/>
        </sheetView>
    </sheetViews>
    <sheetFormatPr defaultRowHeight="15" x14ac:dyDescent="0.25"/>
    <sheetData>
        <row r="1" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A1">
                <v>1</v>
            </c>
        </row>
        <row r="2" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A2">
                <v>2</v>
            </c>
        </row>
        <row r="3" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A3">
                <v>3</v>
            </c>
        </row>
        <row r="4" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A4">
                <v>4</v>
            </c>
        </row>
        <row r="5" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A5">
                <v>5</v>
            </c>
        </row>
        <row r="6" spans="1:1" x14ac:dyDescent="0.25">
            <c r="A6">
                <f>SUM(A1:A5)</f>
                <v>15</v>
            </c>
        </row>
    </sheetData>
    <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>

```



