<!--This is the start of the document-->
# How to: Move a paragraph from one presentation to another (Open XML SDK)
**Last modified:** July 27, 2012

_**Applies to:** Office 2013 | Open XML_

**In this article**

 [Getting a PresentationDocument Object](#sectionSection1)

 [Basic Presentation Document Structure](#sectionSection2)

 [Structure of the Shape Text Body](#sectionSection3)

 [How the Sample Code Works](#sectionSection4)

 [Sample Code](#sectionSection5)



This topic shows how to use the classes in the Open XML SDK 2.5 for Office to move a paragraph from one presentation to another presentation programmatically.

The following assembly directives are required to compile the code in this topic.


```C#
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using Drawing = DocumentFormat.OpenXml.Drawing;
```




```VisualBasic
Imports System.Linq
Imports DocumentFormat.OpenXml.Presentation
Imports DocumentFormat.OpenXml.Packaging
Imports Drawing = DocumentFormat.OpenXml.Drawing
```



<a name="sectionSection1" />




## Getting a PresentationDocument Object
In the Open XML SDK, the  **PresentationDocument** class represents a presentation document package. To work with a presentation document, first create an instance of the **PresentationDocument** class, and then work with that instance. To create the class instance from the document call the **Open(String, Boolean)** method that uses a file path, and a Boolean value as the second parameter to specify whether a document is editable. To open a document for read/write, specify the value **true** for this parameter as shown in the following **using** statement. In this code, the_file_ parameter is a string that represents the path for the file from which you want to open the document.


```C#
using (PresentationDocument doc = PresentationDocument.Open(file, true))
{
    // Insert other code here.
}

```




```VisualBasic
Using doc As PresentationDocument = PresentationDocument.Open(file, True)
    ' Insert other code here.
End Using
```



The  **using** statement provides a recommended alternative to the typical .Open, .Save, .Close sequence. It ensures that the **Dispose** method (internal method used by the Open XML SDK to clean up resources) is automatically called when the closing brace is reached. The block that follows the **using** statement establishes a scope for the object that is created or named in the **using** statement, in this case_doc_.

<a name="sectionSection2" />




## Basic Presentation Document Structure
The basic document structure of a  **PresentationML** document consists of a number of parts, among which is the main part that contains the presentation definition. The following text from the [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification introduces the overall form of a **PresentationML** package.


> A  **PresentationML** package's main part starts with a presentation root element. That element contains a presentation, which, in turn, refers to a **slide** list, a_slide master_ list, a_notes master_ list, and a_handout master_ list. The slide list refers to all of the slides in the presentation; the slide master list refers to the entire slide masters used in the presentation; the notes master contains information about the formatting of notes pages; and the handout master describes how a handout looks.


> A _handout_ is a printed set of slides that can be provided to an_audience_ for future reference.


> As well as text and graphics, each slide can contain _comments_ and_notes_, can have a _layout_, and can be part of one or more _custom presentations_. (A comment is an annotation intended for the person maintaining the presentation slide deck. A note is a reminder or piece of text intended for the presenter or the audience.)


> Other features that a  **PresentationML** document can include the following:_animation_, _audio_, _video_, and _transitions_ between slides.


> A  **PresentationML** document is not stored as one large body in a single part. Instead, the elements that implement certain groupings of functionality are stored in separate parts. For example, all comments in a document are stored in one comment part while each slide has its own part.


> © ISO/IEC29500: 2008.

This following XML code segment represents a presentation that contains two slides denoted by the ID 267 and 256.


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



Using the Open XML SDK 2.5, you can create document structure and content using strongly-typed classes that correspond to PresentationML elements. You can find these classes in the  **DocumentFormat.OpenXml.Presentation** namespace. The following table lists the class names of the classes that correspond to the **sld**,  **sldLayout**,  **sldMaster**, and  **notesMaster** elements.


****

<table xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><tr><th><p>PresentationML Element</p></th><th><p>Open XML SDK 2.5 Class</p></th><th><p>Description</p></th></tr><tr><td><p>sld</p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.Slide"><span class="nolink">Slide</span></span></p></td><td><p>Presentation Slide. It is the root element of SlidePart.</p></td></tr><tr><td><p>sldLayout</p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.SlideLayout"><span class="nolink">SlideLayout</span></span></p></td><td><p>Slide Layout. It is the root element of SlideLayoutPart.</p></td></tr><tr><td><p>sldMaster </p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.SlideMaster"><span class="nolink">SlideMaster</span></span></p></td><td><p>Slide Master. It is the root element of SlideMasterPart.</p></td></tr><tr><td><p>notesMaster</p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.NotesMaster"><span class="nolink">NotesMaster</span></span></p></td><td><p>Notes Master (or handoutMaster). It is the root element of NotesMasterPart.</p></td></tr></table>
<a name="sectionSection3" />




## Structure of the Shape Text Body
The following text from the  [ISO/IEC 29500](http://go.microsoft.com/fwlink/?LinkId=194337) specification introduces the structure of this element.


> This element specifies the existence of text to be contained within the corresponding shape. All visible text and visible text related properties are contained within this element. There can be multiple paragraphs and within paragraphs multiple runs of text.


> © ISO/IEC29500: 2008.

The following table lists the child elements of the shape text body and the description of each.


****

<table xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><tr><th><p>Child Element</p></th><th><p>Description</p></th></tr><tr><td><p>bodyPr</p></td><td><p>Body Properties</p></td></tr><tr><td><p>lstStyle</p></td><td><p>Text List Styles</p></td></tr><tr><td><p>p</p></td><td><p>Text Paragraphs</p></td></tr></table>
The following XML Schema fragment defines the contents of this element:


```XML
<complexType name="CT_TextBody">
   <sequence>
       <element name="bodyPr" type="CT_TextBodyProperties" minOccurs="1" maxOccurs="1"/>
       <element name="lstStyle" type="CT_TextListStyle" minOccurs="0" maxOccurs="1"/>
       <element name="p" type="CT_TextParagraph" minOccurs="1" maxOccurs="unbounded"/>
   </sequence>
</complexType>

```



<a name="sectionSection4" />




## How the Sample Code Works
The code in this topic consists of two methods,  **MoveParagraphToPresentation** and **GetFirstSlide**. The first method takes two string parameters: one that represents the source file, which contains the paragraph to move, and one that represents the target file, to which the paragraph is moved. The method opens both presentation files and then calls the  **GetFirstSlide** method to get the first slide in each file. It then gets the first **TextBody** shape in each slide and the first paragraph in the source shape. It performs adeep clone of the source paragraph, copying not only the source **Paragraph** object itself, but also everything contained in that object, including its text. It then inserts the cloned paragraph in the target file and removes the source paragraph from the source file, replacing it with a placeholder paragraph. Finally, it saves the modified slides in both presentations.


```C#
// Moves a paragraph range in a TextBody shape in the source document
// to another TextBody shape in the target document.
public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
{
    // Open the source file as read/write.
    using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
    {
        // Open the target file as read/write.
        using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
        {
            // Get the first slide in the source presentation.
            SlidePart slide1 = GetFirstSlide(sourceDoc);

            // Get the first TextBody shape in it.
            TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

            // Get the first paragraph in the TextBody shape.
            // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
            Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

            // Get the first slide in the target presentation.
            SlidePart slide2 = GetFirstSlide(targetDoc);

            // Get the first TextBody shape in it.
            TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

            // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
            // Passing "true" creates a deep clone, which creates a copy of the 
            // Paragraph object and everything directly or indirectly referenced by that object.
            textBody2.Append(p1.CloneNode(true));

            // Remove the source paragraph from the source file.
            textBody1.RemoveChild<Drawing.Paragraph>(p1);

            // Replace the removed paragraph with a placeholder.
            textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

            // Save the slide in the source file.
            slide1.Slide.Save();

            // Save the slide in the target file.
            slide2.Slide.Save();
        }
    }
}

```




```VisualBasic
' Moves a paragraph range in a TextBody shape in the source document
' to another TextBody shape in the target document.
Public Shared Sub MoveParagraphToPresentation(ByVal sourceFile As String, ByVal targetFile As String)
    ' Open the source file as read/write.
    Using sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)
        ' Open the target file as read/write.
        Using targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)
            ' Get the first slide in the source presentation.
            Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

            ' Get the first TextBody shape in it.
            Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody)().First()

            ' Get the first paragraph in the TextBody shape.
            ' Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
            Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph)().First()

            ' Get the first slide in the target presentation.
            Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

            ' Get the first TextBody shape in it.
            Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody)().First()

            ' Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
            ' Passing "true" creates a deep clone, which creates a copy of the 
            ' Paragraph object and everything directly or indirectly referenced by that object.
            textBody2.Append(p1.CloneNode(True))

            ' Remove the source paragraph from the source file.
            textBody1.RemoveChild(Of Drawing.Paragraph)(p1)

            ' Replace the removed paragraph with a placeholder.
            textBody1.AppendChild(Of Drawing.Paragraph)(New Drawing.Paragraph())

            ' Save the slide in the source file.
            slide1.Slide.Save()

            ' Save the slide in the target file.
            slide2.Slide.Save()
        End Using
    End Using
End Sub

```



The  **GetFirstSlide** method takes the **PresentationDocument** object passed in, gets its presentation part, and then gets the ID of the first slide in its slide list. It then gets the relationship ID of the slide, gets the slide part from the relationship ID, and returns the slide part to the calling method.


```C#
// Get the slide part of the first slide in the presentation document.
public static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
{
    // Get relationship ID of the first slide
    PresentationPart part = presentationDocument.PresentationPart;
    SlideId slideId = part.Presentation.SlideIdList.GetFirstChild<SlideId>();
    string relId = slideId.RelationshipId;

    // Get the slide part by the relationship ID.
    SlidePart slidePart = (SlidePart)part.GetPartById(relId);

    return slidePart;
}
```




```VisualBasic
' Get the slide part of the first slide in the presentation document.
Public Shared Function GetFirstSlide(ByVal presentationDocument As PresentationDocument) As SlidePart
    ' Get relationship ID of the first slide
    Dim part As PresentationPart = presentationDocument.PresentationPart
    Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
    Dim relId As String = slideId.RelationshipId

    ' Get the slide part by the relationship ID.
    Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

    Return slidePart
End Function

```



<a name="sectionSection5" />




## Sample Code
By using this sample code, you can move a paragraph from one presentation to another. In your program, you can use the following call to the  **MoveParagraphToPresentation** method to move the first paragraph from the presentation file "Myppt4.pptx" to the presentation file "Myppt12.pptx".


```C#
string sourceFile = @"C:\Users\Public\Documents\Myppt4.pptx";
string targetFile = @"C:\Users\Public\Documents\Myppt12.pptx";
MoveParagraphToPresentation(sourceFile, targetFile);
```




```VisualBasic
Dim sourceFile As String = "C:\Users\Public\Documents\Myppt4.pptx"
Dim targetFile As String = "C:\Users\Public\Documents\Myppt12.pptx"
MoveParagraphToPresentation(sourceFile, targetFile)
```



After you run the program take a look on the content of both the source and the target files to see the moved paragraph.

The following is the complete sample code in both C# and Visual Basic.


```C#
// Moves a paragraph range in a TextBody shape in the source document
// to another TextBody shape in the target document.
public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
{
    // Open the source file as read/write.
    using (PresentationDocument sourceDoc = PresentationDocument.Open(sourceFile, true))
    {
        // Open the target file as read/write.
        using (PresentationDocument targetDoc = PresentationDocument.Open(targetFile, true))
        {
            // Get the first slide in the source presentation.
            SlidePart slide1 = GetFirstSlide(sourceDoc);

            // Get the first TextBody shape in it.
            TextBody textBody1 = slide1.Slide.Descendants<TextBody>().First();

            // Get the first paragraph in the TextBody shape.
            // Note: "Drawing" is the alias of namespace DocumentFormat.OpenXml.Drawing
            Drawing.Paragraph p1 = textBody1.Elements<Drawing.Paragraph>().First();

            // Get the first slide in the target presentation.
            SlidePart slide2 = GetFirstSlide(targetDoc);

            // Get the first TextBody shape in it.
            TextBody textBody2 = slide2.Slide.Descendants<TextBody>().First();

            // Clone the source paragraph and insert the cloned. paragraph into the target TextBody shape.
            // Passing "true" creates a deep clone, which creates a copy of the 
            // Paragraph object and everything directly or indirectly referenced by that object.
            textBody2.Append(p1.CloneNode(true));

            // Remove the source paragraph from the source file.
            textBody1.RemoveChild<Drawing.Paragraph>(p1);

            // Replace the removed paragraph with a placeholder.
            textBody1.AppendChild<Drawing.Paragraph>(new Drawing.Paragraph());

            // Save the slide in the source file.
            slide1.Slide.Save();

            // Save the slide in the target file.
            slide2.Slide.Save();
        }
    }
}

// Get the slide part of the first slide in the presentation document.
public static SlidePart GetFirstSlide(PresentationDocument presentationDocument)
{
    // Get relationship ID of the first slide
    PresentationPart part = presentationDocument.PresentationPart;
    SlideId slideId = part.Presentation.SlideIdList.GetFirstChild<SlideId>();
    string relId = slideId.RelationshipId;

    // Get the slide part by the relationship ID.
    SlidePart slidePart = (SlidePart)part.GetPartById(relId);

    return slidePart;
}

```




```VisualBasic
' Moves a paragraph range in a TextBody shape in the source document
' to another TextBody shape in the target document.
Public Sub MoveParagraphToPresentation(ByVal sourceFile As String, ByVal targetFile As String)

    ' Open the source file.
    Dim sourceDoc As PresentationDocument = PresentationDocument.Open(sourceFile, True)

    ' Open the target file.
    Dim targetDoc As PresentationDocument = PresentationDocument.Open(targetFile, True)

    ' Get the first slide in the source presentation.
    Dim slide1 As SlidePart = GetFirstSlide(sourceDoc)

    ' Get the first TextBody shape in it.
    Dim textBody1 As TextBody = slide1.Slide.Descendants(Of TextBody).First()

    ' Get the first paragraph in the TextBody shape.
    ' Note: Drawing is the alias of the namespace DocumentFormat.OpenXml.Drawing
    Dim p1 As Drawing.Paragraph = textBody1.Elements(Of Drawing.Paragraph).First()

    ' Get the first slide in the target presentation.
    Dim slide2 As SlidePart = GetFirstSlide(targetDoc)

    ' Get the first TextBody shape in it.
    Dim textBody2 As TextBody = slide2.Slide.Descendants(Of TextBody).First()

    ' Clone the source paragraph and insert the cloned paragraph into the target TextBody shape.
    textBody2.Append(p1.CloneNode(True))

    ' Remove the source paragraph from the source file.
    textBody1.RemoveChild(Of Drawing.Paragraph)(p1)

    ' Replace it with an empty one, because a paragraph is required for a TextBody shape.
    textBody1.AppendChild(Of Drawing.Paragraph)(New Drawing.Paragraph())

    ' Save the slide in the source file.
    slide1.Slide.Save()

    ' Save the slide in the target file.
    slide2.Slide.Save()

End Sub
' Get the slide part of the first slide in the presentation document.
Public Function GetFirstSlide(ByVal presentationDoc As PresentationDocument) As SlidePart

    ' Get relationship ID of the first slide.
    Dim part As PresentationPart = presentationDoc.PresentationPart
    Dim slideId As SlideId = part.Presentation.SlideIdList.GetFirstChild(Of SlideId)()
    Dim relId As String = slideId.RelationshipId

    ' Get the slide part by the relationship ID.
    Dim slidePart As SlidePart = CType(part.GetPartById(relId), SlidePart)

    Return slidePart

End Function
```




## See also

#### Other resources


 [Open XML SDK 2.5 class library reference](http://msdn.microsoft.com/library/36c8a76e-ce1b-5959-7e85-5d77db7f46d6(Office.15).aspx)
