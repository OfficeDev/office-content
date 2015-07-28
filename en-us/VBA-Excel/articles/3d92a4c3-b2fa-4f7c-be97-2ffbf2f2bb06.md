
# How to: Fill a Value Down into Blank Cells in a Column

 **Last modified:** July 18, 2011

The following example looks at column A, and if there is a blank cell, sets the value of the blank cell to equal the value of the cell above it. This continues down the entire column, filling a value down into each blank cell.

 **Sample code provided by:**
![MVP Contributor](../images/odc_OfficeTA_33px_MVPContrib.jpg) Tom Urtis, [Atlas Programming Management](http://www.atlaspm.com/) | [About the Contributor](3d92a4c3-b2fa-4f7c-be97-2ffbf2f2bb06.md#AboutContributor)




```
Sub FillCellsFromAbove()
    ' Turn off screen updating to improve performance
    Application.ScreenUpdating = False
    On Error Resume Next
    ' Look in column A
    With Columns(1)
        ' For blank cells, set them to equal the cell above
        .SpecialCells(xlCellTypeBlanks).Formula = "=R[-1]C"
        'Convert the formula to a value
        .Value = .Value
    End With
    Err.Clear
    Application.ScreenUpdating = True
End Sub
```


## About the Contributor
<a name="AboutContributor"> </a>

MVP Tom Urtis is the founder of Atlas Programming Management, a full-service Microsoft Office and Excel business solutions company in Silicon Valley. Tom has over 25 years of experience in business management and developing Microsoft Office applications, and is the co-author of "Holy Macro! It's 2,500 Excel VBA Examples." 

