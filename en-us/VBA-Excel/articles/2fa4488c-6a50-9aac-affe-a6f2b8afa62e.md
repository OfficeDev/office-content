
# ChartGroup.AxisGroup Property (Excel)

 **Last modified:** March 10, 2013

 **In this article**
 [Version Information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)
 [Example](#sectionSection3)


Returns or sets the group for the specified chart. Read/write


## Version Information
<a name="sectionSection0"> </a>

Version Added: Excel 2010 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **AxisGroup**

 _expression_A variable that represents a  ** [ChartGroup](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)** object.


### Return Value

 ** [XlAxisGroup](30e0b817-547f-70f8-6e27-4a14031d1d79.md)**


## Remarks
<a name="sectionSection2"> </a>

For 3-D charts, only  **xlPrimary** is valid.


## Example
<a name="sectionSection3"> </a>

This example deletes the value axis if it is in the secondary group.


```
With myChart.Axes(xlValue) 
 If .AxisGroup = xlSecondary Then .Delete 
End With 

```


## See also
<a name="sectionSection3"> </a>


#### Concepts


 [ChartGroup Object](7eee66c5-04a7-fd86-6e34-4c22ccaf8de0.md)
#### Other resources


 [ChartGroup Object Members](2d31f7af-d639-c8f4-0714-08fc618ec92d.md)
