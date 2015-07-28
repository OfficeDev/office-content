
# AddIns2.Item Property (Excel)

 **Last modified:** March 09, 2015

 **In this article**
 [Version Information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Example](#sectionSection3)


Returns a single object from a collection.


## Version Information
<a name="sectionSection0"> </a>

Version Added: Excel 2010 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **Item**( **_Index_**)

 _expression_A variable that returns a  **AddIns2** object.


### Parameters



|**Name**|**Required/Optional**|**Data Type**|**Description**|
|:-----|:-----|:-----|:-----|
|Index|Required| **Variant**|The name or index number of the object.|

## Example
<a name="sectionSection3"> </a>

This example displays the status of the Analysis ToolPak add-in. Note that the string used as the index to the  **AddIns2** method is the **Title** property of the **AddIn** object.


```
If ThisWorkbook.Application.AddIns2.Item("Analysis ToolPak").Installed = True Then 
 MsgBox "Analysis ToolPak add-in is installed" 
Else 
 MsgBox "Analysis ToolPak add-in is not installed" 
End If
```


## See also
<a name="sectionSection3"> </a>


#### Concepts


 [AddIns2 Object](ca4bff78-8ddb-6bc3-b95a-a06a9f75dd88.md)
#### Other resources


 [AddIns2 Object Members](6f9dfc17-648d-a004-2321-d3ed86cd438f.md)
