
# ColorStops.Creator Property (Excel)

 **Last modified:** March 10, 2013

 **In this article**
 [Version Information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)


Returns a 32-bit integer that indicates the application in which this object was created. Read-only  **Long**


## Version Information
<a name="sectionSection0"> </a>

Version Added: Excel 2007 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **Creator**

 _expression_An expression that returns a  **ColorStops** object.


### Return Value

XlCreator


## Remarks
<a name="sectionSection2"> </a>

If the object was created in Microsoft Excel, this property returns the string XCEL, which is equivalent to the hexadecimal number 5843454C. The Creator property is designed to be used in Microsoft Excel for the Macintosh, where each application has a four-character creator code. For example, Microsoft Excel has the creator code XCEL. 


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ColorStops Object](e138347b-f03c-2f50-bf61-f7f2182c9681.md)
#### Other resources


 [ColorStops Object Members](864479e0-3690-70b8-a062-1b48825e00b8.md)
