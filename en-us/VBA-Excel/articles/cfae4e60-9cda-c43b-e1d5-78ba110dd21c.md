
# CategoryCollection.Application Property (Excel)

 **Last modified:** March 10, 2013

 **In this article**
 [Version information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Example](#sectionSection2)
 [Property value](#sectionSection3)


Returns an  ** [Application](19b73597-5cf9-4f56-8227-b5211f657f6f.md)** object that represents the Microsoft Excel application. Read-only.


## Version information
<a name="sectionSection0"> </a>

Version Added: Excel 2013 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **Application**

 _expression_A variable that represents a  [CategoryCollection Object (Excel)](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md) object.


## Example
<a name="sectionSection2"> </a>

This example displays a message about the application that created  `myObject`.


```
Set myObject = ActiveWorkbook 
If myObject.Application.Value = "Microsoft Excel" Then 
   MsgBox "This is an Excel Application object." 
Else 
   MsgBox "This is not an Excel Application object." 
End If
```


## Property value
<a name="sectionSection3"> </a>

 **APPLICATION**


## See also
<a name="sectionSection3"> </a>


#### Other resources


 [CategoryCollection Object Members](39a6f85c-2219-79df-cbbc-0bcc21a517e8.md)
 [CategoryCollection Object](5fc7e8c2-6fcb-8726-36f8-d4ae8c2c91e1.md)
