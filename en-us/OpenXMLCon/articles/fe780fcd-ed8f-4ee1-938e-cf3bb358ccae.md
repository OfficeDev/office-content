<!--This is the start of the document-->
# Structure of a PresentationML document (Open XML SDK)
**Last modified:** July 27, 2012

_**Applies to:** Office 2013 | Open XML_

**In this article**

 [Important Presentation Parts](#sectionSection0)

 [The Structure of a Minimum Presentation File](#sectionSection1)

 [Generated PresentationML XML Code](#sectionSection2)

 [Typical Presentation Scenario](#TypWBScenario)



The document structure of a PresentationML document consists of the <presentation> (Presentation) element that contains <sldMaster> (Slide Master), <sldLayout> (Slide Layout), <sld > (Slide), and <theme> (Theme) elements that reference the slides in the presentation. (The Theme element is the root element of the DrawingMLTheme part.) These elements are the minimum elements required for a valid presentation document. 

In addition, a presentation document might contain <notes> (Notes Slide), <handoutMaster> (Handout Master), <sp> (Shape), <pic> (Picture), <tbl> (Table), and other slide-related elements. (Table elements are defined in the DrawingML schema.)

Other features that a PresentationML document can contain include the following: animation, audio, video, and transitions between slides.

A PresentationML document is not stored as one large body in a single part. Instead, the elements that implement certain groupings of functionality are stored in separate parts. For example, all comments in a document are stored in one comment part, while each slide has its own part. A separate XML file is created for each slide.

<a name="sectionSection0" />




## Important Presentation Parts
Using the Open XML SDK 2.5, you can create document structure and content that uses strongly-typed classes that correspond to PresentationML elements. You can find these classes in the  **DocumentFormat.OpenXml.Presentation** namespace. The following table lists the class names of the classes that correspond to some of the important presentation elements.


****

<table xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:mtps="http://msdn2.microsoft.com/mtps" xmlns:mshelp="http://msdn.microsoft.com/mshelp" xmlns:ddue="http://ddue.schemas.microsoft.com/authoring/2003/5" xmlns:msxsl="urn:schemas-microsoft-com:xslt"><tr><th><p>Package Part</p></th><th><p>Top Level PresentationML Element</p></th><th><p>Open XML SDK 2.5 Class</p></th><th><p>Description*</p></th></tr><tr><td><p>Presentation</p></td><td><p><presentation></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.Presentation"><span class="nolink">Presentation</span></span></p></td><td><p>The root element for the Presentation part. This element specifies within it fundamental presentation-wide properties.</p></td></tr><tr><td><p>Presentation Properties</p></td><td><p><presentationPr></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.PresentationProperties"><span class="nolink">PresentationProperties</span></span></p></td><td><p>The root element for the Presentation Properties part. This element functions as a parent element within which additional presentation-wide document properties are contained.</p></td></tr><tr><td><p>Slide Master</p></td><td><p><sldMaster></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.SlideMaster"><span class="nolink">SlideMaster</span></span></p></td><td><p>The root element for the Slide Master part. Within a slide master slide are contained all elements that describe the objects and their corresponding formatting for within a presentation slide. For more information, see <span sdata="link"><a href="7dfd78a3-e233-4abd-8c17-1e384780d3ec.htm">Working with slide masters (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Slide Layout</p></td><td><p><sldLayout></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.SlideLayout"><span class="nolink">SlideLayout</span></span></p></td><td><p>The root element for the Slide Layout part. This element specifies the relationship information for each slide layout that is used within the slide master. For more information, see <span sdata="link"><a href="1ec087c3-8b9e-46a9-9c3c-14586908eb0e.htm">Working with slide layouts (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Theme</p></td><td><p><officeStyleSheet></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Drawing.Theme"><span class="nolink">Theme</span></span></p></td><td><p>The root element for the Theme part. This element holds all the different formatting options available to a document through a theme and defines the overall look and feel of the document when themed objects are used within the document.</p></td></tr><tr><td><p>Slide</p></td><td><p><sld></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.Slide"><span class="nolink">Slide</span></span></p></td><td><p>The root element for the Slide part. This element specifies a slide within a slide list. For more information, see <span sdata="link"><a href="ee6c905b-26c5-4aed-a414-9aa826364a23.htm">Working with presentation slides (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Notes Master</p></td><td><p><notesMaster></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.NotesMaster"><span class="nolink">NotesMaster</span></span></p></td><td><p>The root element for the Notes Master part. Within a notes master slide are contained all elements that describe the objects and their corresponding formatting for within a notes slide.</p></td></tr><tr><td><p>Notes Slide</p></td><td><p><notes></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.NotesSlide"><span class="nolink">NotesSlide</span></span></p></td><td><p>The root element of the Notes Slide part. This element specifies the existence of a notes slide along with its corresponding data. Contained within a notes slide are all the common slide elements along with addition properties that are specific to the notes element. For more information, see <span sdata="link"><a href="56d28bc5-c9ea-4c0e-b2f5-20be9c16d290.htm">Working with notes slides (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Handout Master</p></td><td><p><handoutMaster></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.HandoutMaster"><span class="nolink">HandoutMaster</span></span></p></td><td><p>The root element of the Handout Master part. Within a handout master slide are contained all elements that describe the objects and their corresponding formatting for within a handout slide. For more information, see <span sdata="link"><a href="fb4b293c-9a23-44b7-8af6-afe5fac6611a.htm">Working with handout master slides (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Comments</p></td><td><p><cmLst></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.CommentList"><span class="nolink">CommentList</span></span></p></td><td><p>The root element of the Comments part. This element specifies a list of comments for a particular slide. For more information, see <span sdata="link"><a href="d7f0f1d3-bcf9-40b5-aaa4-4a08d862ac8e.htm">Working with comments (Open XML SDK)</a></span>.</p></td></tr><tr><td><p>Comments Author</p></td><td><p><cmAuthorLst></p></td><td><p><span sdata="cer" target="T:DocumentFormat.OpenXml.Presentation.CommentAuthorList"><span class="nolink">CommentAuthorList</span></span></p></td><td><p>The root element of the Comments Author part. This element specifies a list of authors with comments in the current document. For more information, see <span sdata="link"><a href="d7f0f1d3-bcf9-40b5-aaa4-4a08d862ac8e.htm">Working with comments (Open XML SDK)</a></span>.</p></td></tr></table>
* Descriptions adapted from the  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification, © ISO/IEC29500: 2008.


### Presentation Part
A PresentationML package's main part starts with a <presentation> root element. That element contains a presentation, which, in turn, refers to a slide list, a slide master list, a notes master list, and a handout master list. The slide list refers to all of the slides in the presentation. The slide master list refers to the entire set of slide masters used in the presentation. The notes master contains information about the formatting of notes pages. The handout master describes how a handout looks. (A handout is a printed set of slides that can be handed out to an audience for future reference.)


### Presentation Properties Part
The root element of the Presentation Properties part is the <presentationPr> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Presentation Properties part as follows:

An instance of this part type contains all the presentation's properties. 

A package shall contain exactly one Presentation Properties part, and that part shall be the target of an implicit relationship from the Presentation (§13.3.6) part. 

[Example: The following Presentation part-relationship item contains a relationship to the Presentation Properties part, which is stored in the ZIP item presProps.xml: 

<Relationships xmlns="…">

 <Relationship Id="rId6"

 Type="http://…/presProps" Target="presProps.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be presentationPr. 

[Example:

<p:presentationPr xmlns:p="…" …>

 <p:clrMru>

 …

 </p:clrMru>

 …

</p:presentationPr>

end example] 

A Presentation Properties part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 

A Presentation Properties part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Slide Master Part
The root element of the Slide Master part is the <sldMaster> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Slide Master part as follows:

An instance of this part type contains the master definition of formatting, text, and objects that appear on each slide in the presentation that is derived from this slide master. 

A package shall contain one or more Slide Master parts, each of which shall be the target of an explicit relationship from the Presentation (§13.3.6) part, as well as an implicit relationship from any Slide Layout (§13.3.9) part where that slide layout is defined based on this slide master. Each can optionally be the target of a relationship in a Slide Layout (§13.3.9) part as well. 

[Example: The following Presentation part-relationship item contains a relationship to the Slide Master part, which is stored in the ZIP item slideMasters/slideMaster1.xml:

<Relationships xmlns="…">

 <Relationship Id="rId1"

 Type="http://…/slideMaster" Target="slideMasters/slideMaster1.xml"/>

</Relationships>

end example] 

The root element for a part of this content type shall be sldMaster. 

[Example: 

<p:sldMaster xmlns:p="…">

 <p:cSld name="">

 …

 </p:cSld>

 <p:clrMap … />

</p:sldMaster>

end example] 

A Slide Master part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 

A Slide Master part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Custom XML Data Storage (§15.2.4) 

• Theme (§14.2.7) 

• Thumbnail (§15.2.16)

A Slide Master part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500:

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• Slide Layout (§13.3.9) 

• Video (§15.2.15)

A Slide Master part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Slide Layout Part
The root element of the Slide Layout part is the <sldLayout> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Slide Layout part as follows:

An instance of this part type contains the definition for a slide layout template for this presentation. This template defines the default appearance and positioning of drawing objects on this slide type when it is created. 

A package shall contain one or more Slide Layout parts, and each of those parts shall be the target of an explicit relationship in the Slide Master (§13.3.10) part, as well as an implicit relationship from each of the Slide (§13.3.8) parts associated with this slide layout. 

[Example: The following Slide Master part-relationship item contains relationships to several Slide Layout parts, which are stored in the ZIP items ../slideLayouts/slideLayoutN.xml:

<Relationships xmlns="…">

 <Relationship Id="rId1"

 Type="http://…/slideLayout"

 Target="../slideLayouts/slideLayout1.xml"/>

 <Relationship Id="rId2"

 Type="http://…/slideLayout"

 Target="../slideLayouts/slideLayout2.xml"/>

 <Relationship Id="rId3"

 Type="http://…/slideLayout"

 Target="../slideLayouts/slideLayout3.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be sldLayout.

[Example:

<p:sldLayout xmlns:p="…" matchingName="" type="title" preserve="1">

 <p:cSld name="Title Slide">

 …

 </p:cSld>

 <p:clrMapOvr>

 <a:masterClrMapping/>

 </p:clrMapOvr>

 <p:timing/>

</p:sldLayout>

end example] 

A Slide Layout part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Custom XML Data Storage (§15.2.4) 

• Slide Master (§13.3.10) 

• Theme Override (§14.2.8) 

• Thumbnail (§15.2.16)

A Slide Layout part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500:

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• Video (§15.2.15)

A Slide Layout part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Slide Part
The root element of the Slide part is the <sld> element.

As well as text and graphics, each slide can contain comments and notes, can have a layout, and can be part of one or more custom presentations. A comment is an annotation intended for the person maintaining the presentation slide deck. A note is a reminder or piece of text intended for the presenter or the audience. 

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Slide part as follows:

A Slide part contains the contents of a single slide.

A package shall contain one Slide part per slide, and each of those parts shall be the target of an explicit relationship from the Presentation (§13.3.6) part.

[Example: Consider a PresentationML document having two slides. The corresponding Presentation part relationship item contains two relationships to Slide parts, which are stored in the ZIP items slides/slide1.xml and slides/slide2.xml:

<Relationships xmlns="…">

 <Relationship Id="rId2"

 Type="http://…/slide" Target="slides/slide1.xml"/>

 <Relationship Id="rId3"

 Type="http://…/slide" Target="slides/slide2.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be sld.

[Example: slides/slide1.xml contains:

<p:sld xmlns:p="…">

 <p:cSld name="">

 …

 </p:cSld>

 <p:clrMapOvr>

 …

 </p:clrMapOvr>

 <p:timing>

 <p:tnLst>

 <p:par>

 <p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"/>

 </p:par>

 </p:tnLst>

 </p:timing>

</p:sld>

end example]

A Slide part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 

A Slide part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Comments (§13.3.2) 

• Custom XML Data Storage (§15.2.4) 

• Notes Slide (§13.3.5) 

• Theme Override (§14.2.8) 

• Thumbnail (§15.2.16) 

• Slide Layout (§13.3.9) 

• Slide Synchronization Data (§13.3.11)

A Slide part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500:

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• User Defined Tags (§13.3.12) 

• Video (§15.2.15)

A Slide part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Theme Part
The root element of the Theme part is the <officeStyleSheet> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML DrawingML Theme part as follows:

An instance of this part type contains information about a document's theme, which is a combination of color scheme, font scheme, and format scheme (the latter also being referred to as effects). For a WordprocessingML document, the choice of theme affects the color and style of headings, among other things. For a SpreadsheetML document, the choice of theme affects the color and style of cell contents and charts, among other things. For a PresentationML document, the choice of theme affects the formatting of slides, handouts, and notes via the associated master, among other things.

A WordprocessingML or SpreadsheetML package shall contain zero or one Theme part, which shall be the target of an implicit relationship in a Main Document (§11.3.10) or Workbook (§12.3.23) part. A PresentationML package shall contain zero or one Theme part per Handout Master (§13.3.3), Notes Master (§13.3.4), Slide Master (§13.3.10) or Presentation (§13.3.6) part via an implicit relationship.

[Example: The following WordprocessingML Main Document part-relationship item contains a relationship to the Theme part, which is stored in the ZIP item theme/theme1.xml:

<Relationships xmlns="…">

 <Relationship Id="rId4"

 Type="http://…/theme" Target="theme/theme1.xml"/>

 </Relationships>

end example]

The root element for a part of this content type shall be officeStyleSheet.

[Example: theme1.xml contains the following, where the name attributes of the clrScheme, fontScheme, and fmtScheme elements correspond to the document's color scheme, font scheme, and format scheme, respectively:

<a:officeStyleSheet xmlns:a="…">

 <a:baseStyles>

 <a:clrScheme name="…">

 …

 </a:clrScheme>

 <a:fontScheme name="…">

 …

 </a:fontScheme>

 <a:fmtScheme name="…">

 …

 </a:fmtScheme>

 </a:baseStyles>

 <a:objectDefaults/>

</a:officeStyleSheet>

end example] 



A Theme part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 



A Theme part is permitted to contain explicit relationships to the following parts defined by ISO/IEC 29500:

• Image (§15.2.14)

A Theme part shall not have any implicit or explicit relationships to other parts defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Notes Master Part
The root element of the Notes Master part is the <notesMaster> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Notes Master part as follows:

An instance of this part type contains information about the content and formatting of all notes pages.

A package shall contain at most one Notes Master part, and that part shall be the target of an implicit relationship from the Notes Slide (§13.3.5) part, as well as an explicit relationship from the Presentation (§13.3.6) part. 

[Example: The following Presentation part-relationship item contains a relationship to the Notes Master part, which is stored in the ZIP item notesMasters/notesMaster1.xml:

<Relationships xmlns="…">

 <Relationship Id="rId4"

 Type="http://…/notesMaster" Target="notesMasters/notesMaster1.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be notesMaster.

[Example:

<p:notesMaster xmlns:p="…">

 <p:cSld name="">

 …

 </p:cSld>

 <p:clrMap … />

</p:notesMaster>

end example]

A Notes Master part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 

A Notes Master part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500: 

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Custom XML Data Storage (§15.2.4) 

• Theme (§14.2.7) 

• Thumbnail (§15.2.16)

A Notes Master part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500: 

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• Video (§15.2.15)

The Notes Master part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Notes Slide Part
The root element of the Notes Slide part is the <notes> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Notes Slide part as follows:

An instance of this part type contains the notes for a single slide. 

A package shall contain one Notes Slide part for each slide that contains notes. If they exist, those parts shall each be the target of an implicit relationship from the Slide (§13.3.8) part. 

[Example: The following Slide part-relationship item contains a relationship to a Notes Slide part, which is stored in the ZIP item ../notesSlides/notesSlide1.xml:

<Relationships xmlns="…">

 <Relationship Id="rId3"

 Type="http://…/notesSlide" Target="../notesSlides/notesSlide1.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be notes.

[Example:

<p:notes xmlns:p="…">

 <p:cSld name="">

 …

 </p:cSld>

 <p:clrMapOvr>

 <a:masterClrMapping/>

 </p:clrMapOvr>

</p:notes>

end example]

A Notes Slide part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal). 

A Notes Slide part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Custom XML Data Storage (§15.2.4) 

• Notes Master (§13.3.4) 

• Theme Override (§14.2.8) 

• Thumbnail (§15.2.16)

A Notes Slide part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500:

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• Video (§15.2.15)

The Notes Slide part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Handout Master Part
The root element of the Handout Master part is the <handoutMaster> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Handout Master part as follows:

An instance of this part type contains the look, position, and size of the slides, notes, header and footer text, date, or page number on the presentation's handout.

A package shall contain at most one Handout Master part, and it shall be the target of an explicit relationship from the Presentation (§13.3.6) part.

[Example: The following Presentation part-relationship item contains a relationship to the Handout Master part, which is stored in the ZIP item handoutMasters/handoutMaster1.xml:

<Relationships xmlns="…">

 <Relationship Id="rId5"

 Type="http://…/handoutMaster"

 Target="handoutMasters/handoutMaster1.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be handoutMaster.

[Example:

<p:handoutMaster xmlns:p="…">

 <p:cSld name="">

 …

 </p:cSld>

 <p:clrMap … />

</p:handoutMaster>

end example] 

A Handout Master part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal).

A Handout Master part is permitted to have implicit relationships to the following parts defined by ISO/IEC 29500:

• Additional Characteristics (§15.2.1) 

• Bibliography (§15.2.3) 

• Custom XML Data Storage (§15.2.4) 

• Theme (§14.2.7) 

• Thumbnail (§15.2.16)

A Handout Master part is permitted to have explicit relationships to the following parts defined by ISO/IEC 29500:

• Audio (§15.2.2) 

• Chart (§14.2.1) 

• Content Part (§15.2.4) 

• Diagrams: Diagram Colors (§14.2.3), Diagram Data (§14.2.4), Diagram Layout Definition (§14.2.5), and Diagram Styles (§14.2.6) 

• Embedded Control Persistence (§15.2.9) 

• Embedded Object (§15.2.10) 

• Embedded Package (§15.2.11) 

• Hyperlink (§15.3) 

• Image (§15.2.14) 

• Video (§15.2.15)

A Handout Master part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500.

© ISO/IEC29500: 2008.


### Comments Part
The root element of the Comments part is the <cmLst> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Comments part as follows:

An instance of this part type contains the comments for a single slide. Each comment is tied to its author via an author-ID. Each comment's index number and author-ID combination are unique.

A package shall contain one Comments part for each slide containing one or more comments, and each of those parts shall be the target of an implicit relationship from its corresponding Slide (§13.3.8) part.

[Example: The following Slide part-relationship item contains a relationship to a Comments part, which is stored in the ZIP item ../comments/comment2.xml:

<Relationships xmlns="…">

 <Relationship Id="rId4"

 Type="http://…/comments"

 Target="../comments/comment2.xml"/>

</Relationships>

end example]

The root element for a part of this content type shall be cmLst.

[Example: The Comments part contains three comments, two created by one author, and one created by another, all at the dates and times shown. The index numbers are assigned on a per-author basis, starting at 1 for an author's first comment:

<p:cmLst xmlns:p="…" …>

 <p:cm authorId="0" dt="2005-11-13T17:00:22.071" idx="1">

 <p:pos x="4486" y="1342"/>

 <p:text>Comment text goes here.</p:text>

 </p:cm>

 <p:cm authorId="0" dt="2005-11-13T17:00:34.849" idx="2">

 <p:pos x="3607" y="1867"/>

 <p:text>Another comment's text goes here.</p:text>

 </p:cm>

 <p:cm authorId="1" dt="2005-11-15T00:06:46.919" idx="1">

 <p:pos x="1493" y="2927"/>

 <p:text>comment …</p:text>

 </p:cm>

</p:cmLst>

end example]

A Comments part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal).

A Comments part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.


### Comments Author Part
The root element of the Comments Author part is the <cmAuthorLst> element.

The  [ISO/IEC 29500](http://www.iso.org/iso/iso_catalogue/catalogue_tc/catalogue_detail.htm?csnumber=51463) specification describes the Open XML PresentationML Comments Author part as follows:

An instance of this part type contains information about each author who has added a comment to the document. That information includes the author's name, initials, a unique author-ID, a last-comment-index-used count, and a display color. (The color can be used when displaying comments to distinguish comments from different authors.)

A package shall contain at most one Comment Authors part. If it exists, that part shall be the target of an implicit relationship from the Presentation (§13.3.6) part.

[Example: The following Presentation part relationship item contains a relationship to the Comment Authors part, which is stored in the ZIP item commentAuthors.xml:

<Relationships xmlns="…">

 <Relationship Id="rId8"

 Type="http://…/commentAuthors" Target="commentAuthors.xml"/>

</Relationships>

end example] 

The root element for a part of this content type shall be cmAuthorLst.

[Example: Two people have authored comments in this document: Mary Smith and Peter Jones. Her initials are "mas", her author-ID is 0, and her comments' display color index is 0. Since Mary's last-comment-index-used value is 3, the next comment-index to be used for her is 4. His initials are "pjj", his author-ID is 1, and his comments' display color index is 1. Since Peter's last-comment-index-used value is 1, the next comment-index to be used for him is 2:

<p:cmAuthorLst xmlns:p="…" …>

 <p:cmAuthor id="0" name="Mary Smith" initials="mas" lastIdx="3" clrIdx="0"/>

 <p:cmAuthor id="1" name="Peter Jones" initials="pjj" lastIdx="1" clrIdx="1"/>

</p:cmAuthorLst>

end example]

A Comment Authors part shall be located within the package containing the relationships part (expressed syntactically, the TargetMode attribute of the Relationship element shall be Internal).

A Comment Authors part shall not have implicit or explicit relationships to any other part defined by ISO/IEC 29500. 

© ISO/IEC29500: 2008.

<a name="sectionSection1" />




## The Structure of a Minimum Presentation File
Now that you are familiar with the parts of a PresentationML document, consider how some of these parts are implemented and connected in an actual presentation file. As shown in the article  [How to: Create a presentation document by providing a file name (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/3d4a800e-64f0-4715-919f-a8f7d92a5c37.md), you can use the Open XML API to build up a minimum presentation file, part by part. 

A minimum presentation file consists of a presentation part, represented by the file presentation.xml, as well as a presentation properties part (presProps.xml), a slide master part (slideMaster.xml), a slide layout part (slideLayout.xml), and a theme part (theme.xml). One or more slide parts (slide.xml) are optional.

The packaging structure of a presentation document contains several references between the parts, including some circular references. For example, slide layouts reference slide masters, and slide masters reference slide layouts.

<a name="sectionSection2" />




## Generated PresentationML XML Code
After you run the Open XML SDK 2.5 code to generate a presentation, you can explore the contents of the .zip package to view the PresentationML XML code. To view the .zip package, rename the extension on the minimum presentation from  **.pptx** to **.zip**. Inside the .zip package, there are several parts that make up the minimum presentation.

Figure 1 shows the structure under the  **ppt** folder of the .zip package for a minimum presentation that contains a single slide.


**Figure 1. Minimum presentation folder structure**



![Minimum presentation folder structure]
(images/odc_oxml_ppt_documentstructure_fig01.jpg)

The presentation.xml file contains <sld> (Slide) elements that reference the slides in the presentation. Each slide is associated to the presentation by means of a slide ID and a relationship ID. The  **slideID** is the identifier (ID) used within the package to identify a slide and must be unique within the presentation. The **id** attribute is the relationship ID that identifies the slide part definition associated with a slide. For more information about the slide part, see [Working with presentation slides (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/ee6c905b-26c5-4aed-a414-9aa826364a23.md).

The following XML code is the PresentationML that represents the presentation part of a presentation document that contains a single slide. This code is generated when you run the Open XML SDK 2.5 code to create a minimum presentation


```XML
<?xml version="1.0" encoding="utf-8"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648"
                   r:id="rId1"
                   xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256"
             r:id="rId2"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" />
  </p:sldIdLst>
  <p:sldSz cx="9144000"
           cy="6858000"
           type="screen4x3" />
  <p:notesSz cx="6858000"
             cy="9144000" />
  <p:defaultTextStyle />
</p:presentation>
```



The following XML code is the PresentationML that represents the relationship part of the presentation document. This code is generated when you run the Open XML SDK 2.5 to create a minimum presentation.


```XML
<?xml version="1.0" encoding="utf-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
                Target="/ppt/slides/slide.xml"
                Id="rId2" />
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
                Target="/ppt/slideLayouts/slideMasters/slideMaster.xml"
                Id="rId1" />
  <Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
                Target="/ppt/slideLayouts/slideMasters/theme/theme.xml"
                Id="rId5" />
</Relationships>
```



The following XML code is the PresentationML that represents the slide part of the presentation document. Each slide in a presentation has a slide part associated with it. This code is generated when you run the Open XML SDK 2.5 to create a minimum presentation.


```XML
<?xml version="1.0" encoding="utf-8"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1"
                 name="" />
        <p:cNvGrpSpPr />
        <p:nvPr />
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2"
                   name="Title 1" />
          <p:cNvSpPr>
            <a:spLocks noGrp="1"
                       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
          </p:cNvSpPr>
          <p:nvPr>
            <p:ph />
          </p:nvPr>
        </p:nvSpPr>
        <p:spPr />
        <p:txBody>
          <a:bodyPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
          <a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
          <a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:endParaRPr lang="en-US" />
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr>
    <a:masterClrMapping xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
  </p:clrMapOvr>
</p:sld>

```



<a name="TypWBScenario" />




## Typical Presentation Scenario
A typical presentation does not have a minimum configuration. A typical presentation might contain several slides, each of which references slide layouts and slide masters, and which might contain comments. In addition, a presentation might contain handouts and notes slides, each of which is represented by separate parts. These additional parts are contained within the .zip package of the presentation document.

Figure 2 shows most of the elements that you would find in a typical presentation.


**Figure 2. Elements of a PresentationML file**



![Elements of a PresentationML file]
(images/odc_oxml_ppt_documentstructure_fig02.jpg)


## See also

#### Concepts


 [How to: Create a presentation document by providing a file name (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/3d4a800e-64f0-4715-919f-a8f7d92a5c37.md)

 [Working with presentations (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/82deb499-7479-474d-9d89-c4847e6f3649.md)

 [Working with presentation slides (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/ee6c905b-26c5-4aed-a414-9aa826364a23.md)

 [Working with slide masters (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/7dfd78a3-e233-4abd-8c17-1e384780d3ec.md)

 [Working with slide layouts (Open XML SDK)](https://github.com/jhershey00/OpenXMLTest/blob/master/articles/1ec087c3-8b9e-46a9-9c3c-14586908eb0e.md)
