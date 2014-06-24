<!--This is the start of the document-->
# How to: Open a presentation document for read-only access (Open XML SDK)
**Last modified:** July 27, 2012

_**Applies to:** Office 2013 | Open XML_

**In this article**

 [How to Open a File for Read-Only Access](#sectionSection1)

 [Create an Instance of the PresentationDocument Class](#sectionSection2)

 [Basic Presentation Document Structure](#sectionSection3)

 [How the Sample Code Works](#sectionSection4)

 [Sample Code](#sectionSection5)



This topic describes how to use the classes in the Open XML SDK 2.5 for Office to programmatically open a presentation document for read-only access.

The following assembly directives are required to compile the code in this topic.


```C#
using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Text;
```




```VisualBasic
Imports System
Imports System.Collections.Generic
Imports DocumentFormat.OpenXml.Presentation
Imports A = DocumentFormat.OpenXml.Drawing
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml
Imports System.Text
```



<a name="sectionSection1" />




## How to Open a File for Read-Only Access
You may want to open a presentation document to read the slides. You might want to extract information from a slide, copy a slide to a slide library, or list the titles of the slides. In such cases you want to do so in a way that ensures the document remains unchanged. You can do that by opening the document for read-only access. This How-To topic discusses several ways to programmatically open a read-only presentation document.

<a name="sectionSection2" />




## Create an Instance of the PresentationDocument Class
In the Open XML SDK, the  **PresentationDocument** class represents a presentation document package. To work with a presentation document, first create an instance of the **PresentationDocument** class, and then work with that instance. To create the class instance from the document call one of the **Open** methods. Several Open methods are provided, each with a different signature. The following table contains a subset of the overloads for the **Open** method that you can use to open the package.


****

<table xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><tr><td><p><span class="label">Name</span></p></td><td><p><span class="label">Description</span></p></td></tr><tr><td><p><span sdata="cer" target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.String,System.Boolean)"><span class="nolink">Open(String, Boolean)</span></span></p></td><td><p>Create a new instance of the <span class="keyword">PresentationDocument</span> class from the specified file.</p></td></tr><tr><td><p><span sdata="cer" target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.IO.Stream,System.Boolean)"><span class="nolink">Open(Stream, Boolean)</span></span></p></td><td><p>Create a new instance of the <span class="keyword">PresentationDocument</span> class from the I/O stream. </p></td></tr><tr><td><p><span sdata="cer" target="M:DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(System.IO.Packaging.Package)"><span class="nolink">Open(Package)</span></span></p></td><td><p>Create a new instance of the <span class="keyword">PresentationDocument</span> class from the specified package.</p></td></tr></table>
The previous table includes two  **Open** methods that accept a Boolean value as the second parameter to specify whether a document is editable. To open a document for read-only access, specify the value **false** for this parameter.

For example, you can open the presentation file as read-only and assign it to a  **PresentationDocument** object as shown in the following **using** statement. In this code, thepresentationFile parameter is a string that represents the path of the file from which you want to open the document.


```C#
using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
{
    // Insert other code here.
}

```




```VisualBasic
Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationFile, False)
    ' Insert other code here.
End Using
```



You can also use the second overload of the  **Open** method, in the table above, to create an instance of the **PresentationDocument** class based on an I/O stream. You might use this approach if you have a Microsoft SharePoint Foundation 2010 application that uses stream I/O and you want to use the Open XML SDK 2.5 to work with a document. The following code segment opens a document based on a stream.


```C#
Stream stream = File.Open(strDoc, FileMode.Open);
using (PresentationDocument presentationDocument =
    PresentationDocument.Open(stream, false)) 
{
    // Place other code here.
}
```




```VisualBasic
Dim stream As Stream = File.Open(strDoc, FileMode.Open)
Using presentationDocument As PresentationDocument = PresentationDocument.Open(stream, False)
    ' Other code goes here.
End Using
```



Suppose you have an application that employs the Open XML support in the  **System.IO.Packaging** namespace of the .NET Framework Class Library, and you want to use the Open XML SDK 2.5 to work with a package read-only. The Open XML SDK 2.5 includes a method overload that accepts a **Package** as the only parameter. There is no Boolean parameter to indicate whether the document should be opened for editing. The recommended approach is to open the package as read-only prior to creating the instance of the **PresentationDocument** class. The following code segment performs this operation.


```C#
Package presentationPackage = Package.Open(filepath, FileMode.Open, FileAccess.Read);
using (PresentationDocument presentationDocument =
    PresentationDocument.Open(presentationPackage))
{
    // Other code goes here.
}
```




```VisualBasic
Dim presentationPackage As Package = Package.Open(filepath, FileMode.Open, FileAccess.Read)
Using presentationDocument As PresentationDocument = PresentationDocument.Open(presentationPackage)
    ' Other code goes here.
End Using
```



<a name="sectionSection3" />




## Basic Presentation Document Structure
The basic document structure of a  **PresentationML** document consists of a number of parts, among which the main part is that contains the presentation definition. The following text from the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification introduces the overall form of a **PresentationML** package.


> The main part of a  **PresentationML** package starts with a presentation root element. That element contains a presentation, which, in turn, refers to a **slide** list, aslide master list, anotes master list, and ahandout master list. The slide list refers to all of the slides in the presentation; the slide master list refers to the entire slide masters used in the presentation; the notes master contains information about the formatting of notes pages; and the handout master describes how a handout looks.


> A handout is a printed set of slides that can be provided to anaudience for future reference.


> As well as text and graphics, each slide can contain comments andnotes, can have a layout, and can be part of one or more custom presentations. A comment is an annotation intended for the person maintaining the presentation slide deck. A note is a reminder or piece of text intended for the presenter or the audience.


> Other features that a  **PresentationML** document can include the following:animation, audio, video, and transitions between slides.


> A  **PresentationML** document is not stored as one large body in a single part. Instead, the elements that implement certain groupings of functionality are stored in separate parts. For example, all comments in a document are stored in one comment part while each slide has its own part.


> © ISO/IEC29500: 2008.

This following XML code segment represents a presentation that contains two slides denoted by the IDs 267 and 256. The  **ID** property specifies the slide identifier that contains a unique value throughout the presentation. The possible values for this attribute are from 256 through 2147483647.


```XML
<p:presentation xmlns:p="…" … > 
   <p:sldMasterIdLst>
      <p:sldMasterId
         xmlns:rel="http://…/relationships" rel:id="rId1"/>
   </p:sldMasterIdLst>
   <p:notesMasterIdLst>
      <p:notesMasterId
         xmlns:rel="http://…/relationships" rel:id="rId4"/>
   </p:notesMasterIdLst>
   <p:handoutMasterIdLst>
      <p:handoutMasterId
         xmlns:rel="http://…/relationships" rel:id="rId5"/>
   </p:handoutMasterIdLst>
   <p:sldIdLst>
      <p:sldId id="267"
         xmlns:rel="http://…/relationships" rel:id="rId2"/>
      <p:sldId id="256"
         xmlns:rel="http://…/relationships" rel:id="rId3"/>
   </p:sldIdLst>
       <p:sldSz cx="9144000" cy="6858000"/>
   <p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>
```



<a name="sectionSection4" />




## How the Sample Code Works
In the sample code, after you open the presentation document in the  **using** statement for read-only access, instantiate the **PresentationPart**, and open the slide list. Then you get the relationship ID of the first slide.


```C#
// Get the relationship ID of the first slide.
PresentationPart part = ppt.PresentationPart;
OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
string relId = (slideIds[index] as SlideId).RelationshipId;
```




```VisualBasic
' Get the relationship ID of the first slide.
Dim part As PresentationPart = ppt.PresentationPart
Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
Dim relId As String = (TryCast(slideIds(index), SlideId)).RelationshipId
```



From the relationship ID, relId, you get the slide part, and then the inner text of the slide by building a text string using  **StringBuilder**.


```C#
// Get the slide part from the relationship ID.
SlidePart slide = (SlidePart)part.GetPartById(relId);

// Build a StringBuilder object.
StringBuilder paragraphText = new StringBuilder();

// Get the inner text of the slide.
IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
foreach (A.Text text in texts)
{
    paragraphText.Append(text.Text);
}
sldText = paragraphText.ToString();

```




```VisualBasic
' Get the slide part from the relationship ID.
Dim slide As SlidePart = CType(part.GetPartById(relId), SlidePart)

' Build a StringBuilder object.
Dim paragraphText As New StringBuilder()

' Get the inner text of the slide.
Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
For Each text As A.Text In texts
    paragraphText.Append(text.Text)
Next text
sldText = paragraphText.ToString()
```



The inner text of the slide, which is an  **out** parameter of the **GetSlideIdAndText** method, is passed back to the main method to be displayed.


**Important**  This example displays only the text in the presentation file. Non-text parts, such as shapes or graphics, are not displayed.


  <a name="sectionSection5" />




## Sample Code
The following example opens a presentation file for read-only access and gets the inner text of a slide at a specified index. To call the method GetSlideIdAndText pass in the full path of the presentation document. Also pass in the **out** parametersldText, which will be assigned a value in the method itself, and then you can display its value in the main program. For example, the following call to the  **GetSlideIdAndText** method gets the inner text in the second slide in a presentation file named "Myppt13.pptx".


**Tip**  The most expected exception in this program is the  **ArgumentOutOfRangeException** exception. It could be thrown if, for example, you have a file with two slides, and you wanted to display the text in slide number 4. Therefore, it is best to use a **try** block when you call the **GetSlideIdAndText** method as shown in the following example.


  
```C#
string file = @"C:\Users\Public\Documents\Myppt13.pptx";
string slideText;
int index = 1;
try
{
    GetSlideIdAndText(out slideText, file, index);
    Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText);
}
catch (ArgumentOutOfRangeException exp)
{
    Console.WriteLine(exp.Message);
}
```




```VisualBasic
Dim file As String = "C:\Users\Public\Documents\Myppt13.pptx"
Dim slideText As String = Nothing
Dim index As Integer = 1
Try
    GetSlideIdAndText(slideText, file, index)
    Console.WriteLine("The text in the slide #{0} is: {1}", index + 1, slideText)
Catch exp As ArgumentOutOfRangeException
    Console.WriteLine(exp.Message)
End Try
```



The following is the complete code listing in C# and Visual Basic.


```C#
public static void GetSlideIdAndText(out string sldText, string docName, int index)
{
    using (PresentationDocument ppt = PresentationDocument.Open(docName, false))
    {
        // Get the relationship ID of the first slide.
        PresentationPart part = ppt.PresentationPart;
        OpenXmlElementList slideIds = part.Presentation.SlideIdList.ChildElements;
        string relId = (slideIds[index] as SlideId).RelationshipId;
        relId = (slideIds[index] as SlideId).RelationshipId;

        // Get the slide part from the relationship ID.
        SlidePart slide = (SlidePart)part.GetPartById(relId);

        // Build a StringBuilder object.
        StringBuilder paragraphText = new StringBuilder();

        // Get the inner text of the slide:
        IEnumerable<A.Text> texts = slide.Slide.Descendants<A.Text>();
        foreach (A.Text text in texts)
        {
            paragraphText.Append(text.Text);
        }
        sldText = paragraphText.ToString();
    }
}

```




```VisualBasic
Public Sub GetSlideIdAndText(ByRef sldText As String, ByVal docName As String, ByVal index As Integer)
    Using ppt As PresentationDocument = PresentationDocument.Open(docName, False)
        ' Get the relationship ID of the first slide.
        Dim part As PresentationPart = ppt.PresentationPart
        Dim slideIds As OpenXmlElementList = part.Presentation.SlideIdList.ChildElements
        Dim relId As String = TryCast(slideIds(index), SlideId).RelationshipId
        relId = TryCast(slideIds(index), SlideId).RelationshipId

        ' Get the slide part from the relationship ID.
        Dim slide As SlidePart = DirectCast(part.GetPartById(relId), SlidePart)

        ' Build a StringBuilder object.
        Dim paragraphText As New StringBuilder()

        ' Get the inner text of the slide:
        Dim texts As IEnumerable(Of A.Text) = slide.Slide.Descendants(Of A.Text)()
        For Each text As A.Text In texts
            paragraphText.Append(text.Text)
        Next
        sldText = paragraphText.ToString()
    End Using
End Sub

```




## See also

#### Other resources


 [Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
