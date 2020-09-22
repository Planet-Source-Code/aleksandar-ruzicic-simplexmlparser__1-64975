
 -----------------------------------------

   SimpleXMLParser (SXP)  v0.2

   Written by krckoorascic
   Contact: krckoorascic@gmail.com

 -----------------------------------------

 SimpleXMLParser (SXP) is a small simple and easy to use xml parser/generator.
 As its name says its 'simple', by 'simple' i mean the way how it's structured but
 it also means that its not as powerful as MSXML, for example.
 If you need to handle (read/write) xml files in your application and you don't
 want your application to depend on some libraries (like msxml*.dll) you can include
 this parser in your project and forget those nasty anoying dll-s!

 I built SXP as a part of the skinning engine i'm writting which is not ment to be
 open source (but, hwo knows? i might release the source when it's completed...)
 and from that reason code is not commented and documented as it should be.

 SXP consist of two class modules:
  clsXMLParser - the parser core, reads/writes xml file
  clsXMLNode   - the xml node
 
 Demo project is included with this package and it shows how to use SXP to read xml
 file and to mirror it into TreeView control.

 However, included demo does not shows how to generate/write xml's. (i was sleepy when
 wrote that demo :p)


 How to read xml files using SXP:

 example.xml:
 ---------------------------------------
 <Example>
	<Section id="1">
		<Entry name="foo"/>
		<Entry name="foo2">Some text...</Entry>
	</Section>
	
	<Section id="2">
		<Entry name="blah">
			<Section id="3">
				<Entry name="tezt" value="demo"/>
				<Entry name="sample" value="dummy"/>
			</Section>
		</Entry>
	</Section>
 </Example>
 ---------------------------------------

 ---------------------------------------
 Dim Parser As New clsXMLParser
 Dim Node As clsXMLNode

 Call Parser.Parse("C:\example.xml")

 Set Node = Parser.ParentNode
 ---------------------------------------

 Node now containts whole xml structure readed from file.
 To read 'name' attribute from 'Section' with id=1 you can do like this:
 ---------------------------------------
 MsgBox Node.enumChild(1).enumChild(1).getAttribute("name")
 ---------------------------------------
 
 As you see you must know the exact structure of xml file to get particular data
 thats because SXP is written for skinning engine and it's supposed to parse the
 file by enumerating all child nodes of main node and their children nodes...


 How to write xml file with SXP:

 ---------------------------------------
 Dim Parser As New clsXMLParser
 Dim n1 As New clsXMLNode
 Dim n2 As New clsXMLNode
 Dim n3 As New clsXMLNode

 n1.Name = "Example"
 Call n1.setAttribute("sxp_version", "0.2")
 n2.Name = "Section"
 n3.Name = "Entry"
 n3.Text = "SimpleXMLParser v0.2 by krckoorascic"
 Call n3.setAttribute("name", "description")
 Call n2.addChild(n3)
 Call n1.addChild(n2)

 Parser.ParentNode = n1
 Call Parser.Save("C:\example.xml", True)
 ---------------------------------------

 will generate this file:
 ---------------------------------------
 <Example sxp_version="0.2">
	<Section>
		<Entry name="description">SimpleXMLParser v0.2 by krckoorascic</Entry>
	</Section>
 </Example>
 ---------------------------------------

 NOTE: PrettyPrint parametar of Save method can be used for 'PrettyPrinted' output
 (that means that all nodes are idented, like in example above) or 'single-line'
 output (this is default behavior, it can be useful if you use xml jus for data storage
 which is not ment to be edited by user manualy)
 When setting PrettyPrint to True you can also choose the way how the nodes are idented
 with tabs (default) or with spaces (1 ident=4 spaces).

 More notes:
 SXP is work in progress (but this version is quite stable) so some things are not yet
 done, like better error handling (it will recognize only two types of errors).
 Line and column info in Err.Description is not always accurate...

 About demo project:
 To run demo you need Common Controls 6 and Common Dialog controls.
 studio.xnf.xml is sample xml file on which you can do testing.
 I get that file from winamp :D (orig fname: studio.xnf) cuz i was
 to lazy to write larger xml (or to generate it) to be used for testing.

 
 Copyright n stuff:

 SimpleXMLParser © 2006. Aleksandar Ruzicic (aka krckoorascic). All Rights reserved.

 You're free to use this code in your (free or not free) projects but this code is
 released AS-IS and author is NOT responsable for any possible data lost caused by this code.

 Please report comments, bugs (and possible solutions) to krckoorascic@gmail.com

 Thanks for trying this out!