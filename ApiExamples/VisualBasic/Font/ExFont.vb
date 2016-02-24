' Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Words. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////


Imports Microsoft.VisualBasic
Imports System
Imports System.Collections
Imports System.Drawing
Imports Aspose.Words
Imports Aspose.Words.Drawing
Imports Aspose.Words.Fields
Imports Aspose.Words.Fonts
Imports Aspose.Words.Tables
Imports NUnit.Framework


Namespace ApiExamples.Font
	<TestFixture> _
	Public Class ExFont
		Inherits ApiExampleBase
		<Test> _
		Public Sub CreateFormattedRun()
			'ExStart
			'ExFor:Document.#ctor
			'ExFor:Font
			'ExFor:Font.Name
			'ExFor:Font.Size
			'ExFor:Font.HighlightColor
			'ExFor:Run
			'ExFor:Run.#ctor(DocumentBase,String)
			'ExFor:Story.FirstParagraph
			'ExSummary:Shows how to add a formatted run of text to a document using the object model.
			' Create an empty document. It contains one empty paragraph.
			Dim doc As New Aspose.Words.Document()

			' Create a new run of text.
			Dim run As New Run(doc, "Hello")

			' Specify character formatting for the run of text.
			Dim f As Aspose.Words.Font = run.Font
			f.Name = "Courier New"
			f.Size = 36
			f.HighlightColor = Color.Yellow

			' Append the run of text to the end of the first paragraph
			' in the body of the first section of the document.
			doc.FirstSection.Body.FirstParagraph.AppendChild(run)
			'ExEnd
		End Sub

		<Test> _
		Public Sub Caps()
			'ExStart
			'ExFor:Font.AllCaps
			'ExFor:Font.SmallCaps
			'ExSummary:Shows how to use all capitals and small capitals character formatting properties.
			' Create an empty document. It contains one empty paragraph.
			Dim doc As New Aspose.Words.Document()

			' Get the paragraph from the document, we will be adding runs of text to it.
			Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

			Dim run As New Run(doc, "All capitals")
			run.Font.AllCaps = True
			para.AppendChild(run)

			run = New Run(doc, "SMALL CAPITALS")
			run.Font.SmallCaps = True
			para.AppendChild(run)
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetDocumentFonts()
			'ExStart:
			'ExFor:FontInfoCollection
			'ExFor:DocumentBase.FontInfos
			'ExFor:FontInfo
			'ExFor:FontInfo.Name
			'ExFor:FontInfo.IsTrueType
			'ExSummary:Shows how to gather the details of what fonts are present in a document.
			Dim doc As New Aspose.Words.Document(MyDir & "Document.doc")

			Dim fonts As FontInfoCollection = doc.FontInfos
			Dim fontIndex As Integer = 1

			' The fonts info extracted from this document does not necessarily mean that the fonts themselves are
			' used in the document. If a font is present but not used then most likely they were referenced at some time
			' and then removed from the Document.
			For Each info As FontInfo In fonts
				' Print out some important details about the font.
				Console.WriteLine("Font #{0}", fontIndex)
				Console.WriteLine("Name: {0}", info.Name)
				Console.WriteLine("IsTrueType: {0}", info.IsTrueType)
				fontIndex += 1
			Next info
			'ExEnd
		End Sub

		<Test> _
		Public Sub Strikethrough()
			'ExStart
			'ExFor:Font.StrikeThrough
			'ExFor:Font.DoubleStrikeThrough
			'ExSummary:Shows how to use strike-through character formatting properties.
			' Create an empty document. It contains one empty paragraph.
			Dim doc As New Aspose.Words.Document()

			' Get the paragraph from the document, we will be adding runs of text to it.
			Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

			Dim run As New Run(doc, "Double strike through text")
			run.Font.DoubleStrikeThrough = True
			para.AppendChild(run)

			run = New Run(doc, "Single strike through text")
			run.Font.StrikeThrough = True
			para.AppendChild(run)
			'ExEnd
		End Sub

		<Test> _
		Public Sub PositionSubscript()
			'ExStart
			'ExFor:Font.Position
			'ExFor:Font.Subscript
			'ExFor:Font.Superscript
			'ExSummary:Shows how to use subscript, superscript and baseline text position properties.
			' Create an empty document. It contains one empty paragraph.
			Dim doc As New Aspose.Words.Document()

			' Get the paragraph from the document, we will be adding runs of text to it.
			Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

			' Add a run of text that is raised 5 points above the baseline.
			Dim run As New Run(doc, "Raised text")
			run.Font.Position = 5
			para.AppendChild(run)

			' Add a run of normal text.
			run = New Run(doc, "Normal text")
			para.AppendChild(run)

			' Add a run of text that appears as subscript.
			run = New Run(doc, "Subscript")
			run.Font.Subscript = True
			para.AppendChild(run)

			' Add a run of text that appears as superscript.
			run = New Run(doc, "Superscript")
			run.Font.Superscript = True
			para.AppendChild(run)
			'ExEnd
		End Sub

		<Test> _
		Public Sub ScalingSpacing()
			'ExStart
			'ExFor:Font.Scaling
			'ExFor:Font.Spacing
			'ExSummary:Shows how to use character scaling and spacing properties.
			' Create an empty document. It contains one empty paragraph.
			Dim doc As New Aspose.Words.Document()

			' Get the paragraph from the document, we will be adding runs of text to it.
			Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 0, True), Paragraph)

			' Add a run of text with characters 150% width of normal characters.
			Dim run As New Run(doc, "Wide characters")
			run.Font.Scaling = 150
			para.AppendChild(run)

			' Add a run of text with extra 1pt space between characters.
			run = New Run(doc, "Expanded by 1pt")
			run.Font.Spacing = 1
			para.AppendChild(run)

			' Add a run of text with with space between characters reduced by 1pt.
			run = New Run(doc, "Condensed by 1pt")
			run.Font.Spacing = -1
			para.AppendChild(run)
			'ExEnd
		End Sub

		<Test> _
		Public Sub EmbossItalic()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Emboss
			'ExFor:Font.Italic
			'ExSummary:Shows how to create a run of formatted text.
			Dim run As New Run(doc, "Hello")
			run.Font.Emboss = True
			run.Font.Italic = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub Engrave()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Engrave
			'ExSummary:Shows how to create a run of text formatted as engraved.
			Dim run As New Run(doc, "Hello")
			run.Font.Engrave = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub Shadow()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Shadow
			'ExSummary:Shows how to create a run of text formatted with a shadow.
			Dim run As New Run(doc, "Hello")
			run.Font.Engrave = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub Outline()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Outline
			'ExSummary:Shows how to create a run of text formatted as outline.
			Dim run As New Run(doc, "Hello")
			run.Font.Outline = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub Hidden()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Hidden
			'ExSummary:Shows how to create a hidden run of text.
			Dim run As New Run(doc, "Hello")
			run.Font.Hidden = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub Kerning()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Kerning
			'ExSummary:Shows how to specify the font size at which kerning starts.
			Dim run As New Run(doc, "Hello")
			run.Font.Kerning = 24
			'ExEnd
		End Sub

		<Test> _
		Public Sub NoProofing()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.NoProofing
			'ExSummary:Shows how to specify that the run of text is not to be spell checked by Microsoft Word.
			Dim run As New Run(doc, "Hello")
			run.Font.NoProofing = True
			'ExEnd
		End Sub

		<Test> _
		Public Sub LocaleId()
			Dim doc As New Aspose.Words.Document()

			'ExStart
			'ExFor:Font.LocaleId
			'ExSummary:Shows how to specify the language of a text run so Microsoft Word can use a proper spell checker.
			'Create a run of text that contains Russian text.
			Dim run As New Run(doc, "������")

			'Specify the locale so Microsoft Word recognizes this text as Russian.
			'For the list of locale identifiers see http://www.microsoft.com/globaldev/reference/lcid-all.mspx
			run.Font.LocaleId = 1049
			'ExEnd
		End Sub

		<Test> _
		Public Sub Underlines()
			Dim doc As New Aspose.Words.Document()
			'ExStart
			'ExFor:Font.Underline
			'ExFor:Font.UnderlineColor
			'ExSummary:Shows how use the underline character formatting properties.
			Dim run As New Run(doc, "Hello")
			run.Font.Underline = Underline.Dotted
			run.Font.UnderlineColor = Color.Red
			'ExEnd
		End Sub

		<Test> _
		Public Sub Shading()
			'ExStart
			'ExFor:Font.Shading
			'ExSummary:Shows how to apply shading for a run of text.
			Dim builder As New DocumentBuilder()

			Dim shd As Shading = builder.Font.Shading
			shd.Texture = TextureIndex.TextureDiagonalCross
			shd.BackgroundPatternColor = Color.Blue
			shd.ForegroundPatternColor = Color.BlueViolet

			builder.Font.Color = Color.White

			builder.Writeln("White text on a blue background with texture.")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Bidi()
			'ExStart
			'ExFor:Font.Bidi
			'ExFor:Font.NameBi
			'ExFor:Font.SizeBi
			'ExFor:Font.ItalicBi
			'ExFor:Font.BoldBi
			'ExFor:Font.LocaleIdBi
			'ExSummary:Shows how to insert and format right-to-left text.
			Dim builder As New DocumentBuilder()

			' Signal to Microsoft Word that this run of text contains right-to-left text.
			builder.Font.Bidi = True

			' Specify the font and font size to be used for the right-to-left text.
			builder.Font.NameBi = "Andalus"
			builder.Font.SizeBi = 48

			' Specify that the right-to-left text in this run is bold and italic.
			builder.Font.ItalicBi = True
			builder.Font.BoldBi = True

			' Specify the locale so Microsoft Word recognizes this text as Arabic - Saudi Arabia.
			' For the list of locale identifiers see http://www.microsoft.com/globaldev/reference/lcid-all.mspx
			builder.Font.LocaleIdBi = 1025

			' Insert some Arabic text.
			builder.Writeln("??????")

			builder.Document.Save(MyDir & "Font.Bidi Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub FarEast()
			'ExStart
			'ExFor:Font.NameFarEast
			'ExFor:Font.LocaleIdFarEast
			'ExSummary:Shows how to insert and format text in Chinese or any other Far East language.
			Dim builder As New DocumentBuilder()

			builder.Font.Size = 48

			' Specify the font name. Make sure it the font has the glyphs that you want to display.
			builder.Font.NameFarEast = "SimSun"

			' Specify the locale so Microsoft Word recognizes this text as Chinese.
			' For the list of locale identifiers see http://www.microsoft.com/globaldev/reference/lcid-all.mspx
			builder.Font.LocaleIdFarEast = 2052

			' Insert some Chinese text.
			builder.Writeln("????")

			builder.Document.Save(MyDir & "Font.FarEast Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Names()
			'ExStart
			'ExFor:Font.NameAscii
			'ExFor:Font.NameOther
			'ExSummary:A pretty unusual example of how Microsoft Word can combine two different fonts in one run.
			Dim builder As New DocumentBuilder()

			' This tells Microsoft Word to use Arial for characters 0..127 and
			' Times New Roman for characters 128..255. 
			' Looks like a pretty strange case to me, but it is possible.
			builder.Font.NameAscii = "Arial"
			builder.Font.NameOther = "Times New Roman"

			builder.Writeln("Hello, ������")

			builder.Document.Save(MyDir & "Font.Names Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeStyleIdentifier()
			'ExStart
			'ExFor:Font.StyleIdentifier
			'ExFor:StyleIdentifier
			'ExSummary:Shows how to use style identifier to find text formatted with a specific character style and apply different character style.
			Dim doc As New Aspose.Words.Document(MyDir & "Font.StyleIdentifier.doc")

			' Select all run nodes in the document.
			Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)

			' Loop through every run node.
			For Each run As Run In runs
				' If the character style of the run is what we want, do what we need. Change the style in this case.
				' Note that using StyleIdentifier we can identify a built-in style regardless 
				' of the language of Microsoft Word used to create the document.
				If run.Font.StyleIdentifier.Equals(StyleIdentifier.Emphasis) Then
					run.Font.StyleIdentifier = StyleIdentifier.Strong
				End If
			Next run

			doc.Save(MyDir & "Font.StyleIdentifier Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub ChangeStyleName()
			'ExStart
			'ExFor:Font.StyleName
			'ExSummary:Shows how to use style name to find text formatted with a specific character style and apply different character style.
			Dim doc As New Aspose.Words.Document(MyDir & "Font.StyleName.doc")

			' Select all run nodes in the document.
			Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)

			' Loop through every run node.
			For Each run As Run In runs
				' If the character style of the run is what we want, do what we need. Change the style in this case.
				' Note that names of built in styles could be different in documents 
				' created by Microsoft Word versions for different languages.
				If run.Font.StyleName.Equals("Emphasis") Then
					run.Font.StyleName = "Strong"
				End If
			Next run

			doc.Save(MyDir & "Font.StyleName Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub Style()
			'ExStart
			'ExFor:Font.Style
			'ExFor:Style.BuiltIn
			'ExSummary:Applies double underline to all runs in a document that are formatted with custom character styles.
			Dim doc As New Aspose.Words.Document(MyDir & "Font.Style.doc")

			' Select all run nodes in the document.
			Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)

			' Loop through every run node.
			For Each run As Run In runs
				Dim charStyle As Aspose.Words.Style = run.Font.Style

				' If the style of the run is not a built-in character style, apply double underline.
				If (Not charStyle.BuiltIn) Then
					run.Font.Underline = Underline.Double
				End If
			Next run

			doc.Save(MyDir & "Font.Style Out.doc")
			'ExEnd
		End Sub

		<Test> _
		Public Sub GetAllFonts()
			'ExStart
			'ExFor:Run
			'ExSummary:Gets all fonts used in a document.
			Dim doc As New Aspose.Words.Document(MyDir & "Font.Names.doc")

			' Select all runs in the document.
			Dim runs As NodeCollection = doc.GetChildNodes(NodeType.Run, True)

			' Use a hashtable so we will keep only unique font names.
			Dim fontNames As New Hashtable()

			For Each run As Run In runs
				' This adds an entry into the hashtable.
				' The key is the font name. The value is null, we don't need the value.
				fontNames(run.Font.Name) = Nothing
			Next run

			' There are two fonts used in this document.
			Console.WriteLine("Font Count: " & fontNames.Count)
			'ExEnd

			' Verify the font count is correct.
			Assert.AreEqual(2, fontNames.Count)

		End Sub

		''' <summary>
		''' This calls the below method to resolve skipping of [Test] in VB.NET.
		''' </summary>
		<Test> _
		Public Sub RemoveHiddenContentCaller()
			RemoveHiddenContentFromDocument()
		End Sub

		'ExStart
		'ExFor:Font.Hidden
		'ExFor:Paragraph.Accept
		'ExFor:DocumentVisitor.VisitParagraphStart(Aspose.Words.Paragraph)
		'ExFor:DocumentVisitor.VisitFormField(Aspose.Words.Fields.FormField)
		'ExFor:DocumentVisitor.VisitTableEnd(Aspose.Words.Tables.Table)
		'ExFor:DocumentVisitor.VisitCellEnd(Aspose.Words.Tables.Cell)
		'ExFor:DocumentVisitor.VisitRowEnd(Aspose.Words.Tables.Row)
		'ExFor:DocumentVisitor.VisitSpecialChar(Aspose.Words.SpecialChar)
		'ExFor:DocumentVisitor.VisitGroupShapeStart(Aspose.Words.Drawing.GroupShape)
		'ExFor:DocumentVisitor.VisitShapeStart(Aspose.Words.Drawing.Shape)
		'ExFor:DocumentVisitor.VisitCommentStart(Aspose.Words.Comment)
		'ExFor:DocumentVisitor.VisitFootnoteStart(Aspose.Words.Footnote)
		'ExFor:SpecialChar
		'ExFor:Node.Accept
		'ExFor:Paragraph.ParagraphBreakFont
		'ExFor:Table.Accept
		'ExSummary:Implements the Visitor Pattern to remove all content formatted as hidden from the document.
		Public Sub RemoveHiddenContentFromDocument()
			' Open the document we want to remove hidden content from.
			Dim doc As New Aspose.Words.Document(MyDir & "Font.Hidden.doc")

			' Create an object that inherits from the DocumentVisitor class.
			Dim hiddenContentRemover As New RemoveHiddenContentVisitor()

			' This is the well known Visitor pattern. Get the model to accept a visitor.
			' The model will iterate through itself by calling the corresponding methods
			' on the visitor object (this is called visiting).

			' We can run it over the entire the document like so:
			doc.Accept(hiddenContentRemover)

			' Or we can run it on only a specific node.
			Dim para As Paragraph = CType(doc.GetChild(NodeType.Paragraph, 4, True), Paragraph)
			para.Accept(hiddenContentRemover)

			' Or over a different type of node like below.
			Dim table As Table = CType(doc.GetChild(NodeType.Table, 0, True), Table)
			table.Accept(hiddenContentRemover)

			doc.Save(MyDir & "Font.Hidden Out.doc")

			Assert.AreEqual(13, doc.GetChildNodes(NodeType.Paragraph, True).Count) 'ExSkip
			Assert.AreEqual(1, doc.GetChildNodes(NodeType.Table, True).Count) 'ExSkip
		End Sub

		''' <summary>
		''' This class when executed will remove all hidden content from the Document. Implemented as a Visitor.
		''' </summary>
		Private Class RemoveHiddenContentVisitor
			Inherits DocumentVisitor
			''' <summary>
			''' Called when a FieldStart node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldStart(ByVal fieldStart As FieldStart) As VisitorAction
				' If this node is hidden, then remove it.
				If isHidden(fieldStart) Then
					fieldStart.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FieldEnd node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldEnd(ByVal fieldEnd As FieldEnd) As VisitorAction
				If isHidden(fieldEnd) Then
					fieldEnd.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FieldSeparator node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFieldSeparator(ByVal fieldSeparator As FieldSeparator) As VisitorAction
				If isHidden(fieldSeparator) Then
					fieldSeparator.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a Run node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitRun(ByVal run As Run) As VisitorAction
				If isHidden(run) Then
					run.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a Paragraph node is encountered in the document.
			''' </summary>
			Public Overrides Function VisitParagraphStart(ByVal paragraph As Paragraph) As VisitorAction
				If isHidden(paragraph) Then
					paragraph.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a FormField is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFormField(ByVal field As FormField) As VisitorAction
				If isHidden(field) Then
					field.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a GroupShape is encountered in the document.
			''' </summary>
			Public Overrides Function VisitGroupShapeStart(ByVal groupShape As GroupShape) As VisitorAction
				If isHidden(groupShape) Then
					groupShape.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a Shape is encountered in the document.
			''' </summary>
			Public Overrides Function VisitShapeStart(ByVal shape As Shape) As VisitorAction
				If isHidden(shape) Then
					shape.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a Comment is encountered in the document.
			''' </summary>
			Public Overrides Function VisitCommentStart(ByVal comment As Aspose.Words.Comment) As VisitorAction
				If isHidden(comment) Then
					comment.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a Footnote is encountered in the document.
			''' </summary>
			Public Overrides Function VisitFootnoteStart(ByVal footnote As Footnote) As VisitorAction
				If isHidden(footnote) Then
					footnote.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when visiting of a Table node is ended in the document.
			''' </summary>
			Public Overrides Function VisitTableEnd(ByVal table As Table) As VisitorAction
				' At the moment there is no way to tell if a particular Table/Row/Cell is hidden. 
				' Instead, if the content of a table is hidden, then all inline child nodes of the table should be 
				' hidden and thus removed by previous visits as well. This will result in the container being empty
				' so if this is the case we know to remove the table node.
				'
				' Note that a table which is not hidden but simply has no content will not be affected by this algorthim,
				' as technically they are not completely empty (for example a properly formed Cell will have at least 
				' an empty paragraph in it)
				If (Not table.HasChildNodes) Then
					table.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when visiting of a Cell node is ended in the document.
			''' </summary>
			Public Overrides Function VisitCellEnd(ByVal cell As Cell) As VisitorAction
				If (Not cell.HasChildNodes) AndAlso cell.ParentNode IsNot Nothing Then
					cell.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when visiting of a Row node is ended in the document.
			''' </summary>
			Public Overrides Function VisitRowEnd(ByVal row As Row) As VisitorAction
				If (Not row.HasChildNodes) AndAlso row.ParentNode IsNot Nothing Then
					row.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Called when a SpecialCharacter is encountered in the document.
			''' </summary>
			Public Overrides Function VisitSpecialChar(ByVal character As SpecialChar) As VisitorAction
				If isHidden(character) Then
					character.Remove()
				End If

				Return VisitorAction.Continue
			End Function

			''' <summary>
			''' Returns true if the node passed is set as hidden, returns false if it is visible.
			''' </summary>
			Private Function isHidden(ByVal node As Aspose.Words.Node) As Boolean
				If TypeOf node Is Aspose.Words.Inline Then
					' If the node is Inline then cast it to retrieve the Font property which contains the hidden property
					Dim currentNode As Aspose.Words.Inline = CType(node, Aspose.Words.Inline)
					Return currentNode.Font.Hidden
				ElseIf node.NodeType = NodeType.Paragraph Then
					' If the node is a paragraph cast it to retrieve the ParagraphBreakFont which contains the hidden property
					Dim para As Paragraph = CType(node, Paragraph)
					Return para.ParagraphBreakFont.Hidden
				ElseIf TypeOf node Is ShapeBase Then
					' Node is a shape or groupshape.
					Dim shape As ShapeBase = CType(node, ShapeBase)
					Return shape.Font.Hidden
				ElseIf TypeOf node Is InlineStory Then
					' Node is a comment or footnote.
					Dim inlineStory As InlineStory = CType(node, InlineStory)
					Return inlineStory.Font.Hidden
				End If

				' A node that is passed to this method which does not contain a hidden property will end up here. 
				' By default nodes are not hidden so return false.
				Return False
			End Function

		End Class
		'ExEnd
	End Class
End Namespace
