
# ODBCConnection.SourceDataFile Property (Excel)

 **Last modified:** March 10, 2013

 **In this article**
 [Version Information](#sectionSection0)
 [Syntax](#sectionSection1)
 [Remarks](#sectionSection2)


Returns or sets a  **String** indicating the source data file for an ODBC connection. Read/write.


## Version Information
<a name="sectionSection0"> </a>

Version Added: Excel 2007 


## Syntax
<a name="sectionSection1"> </a>

 _expression_. **SourceDataFile**

 _expression_A variable that represents an  **ODBCConnection** object.


## Remarks
<a name="sectionSection2"> </a>

For file-based data sources (for example, Access) the  **SourceDataFile** property contains a fully qualified path to the source data file. It is null for server-based data sources (for example, SQL Server). The **SourceDataFile** property is set to null if the ** [Connection](2fcd1043-b088-cfde-9853-4a20da20be26.md)** property is changed programmatically.


## See also
<a name="sectionSection2"> </a>


#### Concepts


 [ODBCConnection Object](b880ebec-15a4-5a3d-ef02-db73106db9c9.md)
#### Other resources


 [ODBCConnection Object Members](d13b91f3-a89f-7dd7-7a98-f1d952f3b047.md)
