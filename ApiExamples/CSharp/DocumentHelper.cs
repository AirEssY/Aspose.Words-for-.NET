﻿using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace ApiExamples
{
    /// <summary>
    /// Functions for operations with document and content
    /// </summary>
    internal static class DocumentHelper
    {
        /// <summary>
        /// Create new document without run in the paragraph
        /// </summary>
        internal static Aspose.Words.Document CreateDocumentWithoutDummyText()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            //Create paragraph without run
            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln();

            return doc;
        }

        /// <summary>
        /// Create new document with text
        /// </summary>
        internal static Aspose.Words.Document CreateDocumentFillWithDummyText()
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();

            //Remove the previous changes of the document
            doc.RemoveAllChildren();

            //Set the document author
            doc.BuiltInDocumentProperties.Author = "Test Author";

            DocumentBuilder builder = new DocumentBuilder(doc);

            //Insert new table with two rows and two cells
            InsertTable(doc);

            builder.Writeln("Hello World!");

            // Continued on page 2 of the document content
            builder.InsertBreak(BreakType.PageBreak);

            //Insert TOC entries
            InsertToc(doc);

            return doc;
        }

        /// <summary>
        /// Create new document with textbox shape and some query
        /// </summary>
        internal static Aspose.Words.Document CreateTemplateDocumentForReportingEngine(string templateText)
        {
            Aspose.Words.Document doc = new Aspose.Words.Document();

            //ToDo: Maybe in future add shape(object) as parameter
            // Create textbox shape.
            Shape textbox = new Shape(doc, ShapeType.TextBox);
            textbox.Width = 431.5;
            textbox.Height = 346.35;

            Paragraph paragraph = new Paragraph(doc);
            paragraph.AppendChild(new Run(doc, templateText));

            // Insert paragraph into the textbox.
            textbox.AppendChild(paragraph);

            // Insert textbox into the document.
            doc.FirstSection.Body.FirstParagraph.AppendChild(textbox);

            return doc;
        }


        /// <summary>
        /// Insert new table in the document
        /// </summary>
        private static void InsertTable(Aspose.Words.Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);
            
            //Start creating a new table
            Table table = builder.StartTable();

            //Insert Row 1 Cell 1
            builder.InsertCell();
            builder.Write("Date");

            //Set width to fit the table contents
            table.AutoFit(AutoFitBehavior.AutoFitToContents);
            
            //Insert Row 1 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            //Insert Row 2 Cell 1
            builder.InsertCell();
            builder.Write("Author");

            //Insert Row 2 Cell 2
            builder.InsertCell();
            builder.Write(" ");

            builder.EndRow();

            builder.EndTable();
        }

        /// <summary>
        /// Insert TOC entries in the document
        /// </summary>
        private static void InsertToc(Aspose.Words.Document doc)
        {
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Creating TOC entries
            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading2;

            builder.Writeln("Heading 1.1");
            builder.Writeln("Heading 1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading4;

            builder.Writeln("Heading 1.1.1.1");
            builder.Writeln("Heading 1.1.1.2");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading1;

            builder.Writeln("Heading 2.1");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading5;

            builder.Writeln("Heading 2.1.1.1.1");
            builder.Writeln("Heading 2.1.1.1.2");
            builder.Writeln("Heading 2.1.1.1.3");

            builder.ParagraphFormat.StyleIdentifier = StyleIdentifier.Heading9;

            builder.Writeln("Heading 2.1.1.1.1.1.1.1.1");
            builder.Write("Heading 2.1.1.1.1.1.1.1.2");
        }

        /// <summary>
        /// Insert text into the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="text">
        /// Custom text
        /// </param>
        internal static Run InsertNewRun(Aspose.Words.Document doc, string text)
        {
            Paragraph para = GetParagraph(doc, 0);

            Run run = new Run(doc) { Text = text };

            para.AppendChild(run);

            return run;
        }

        /// <summary>
        /// Get paragraph text of the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="paraIndex">
        /// Paragraph number from collection
        /// </param>
        internal static string GetParagraphText(Aspose.Words.Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex].GetText();
        }

        /// <summary>
        /// Get paragraph of the current document
        /// </summary>
        /// <param name="doc">
        /// Current document
        /// </param>
        /// <param name="paraIndex">
        /// Paragraph number from collection
        /// </param>
        internal static Paragraph GetParagraph(Aspose.Words.Document doc, int paraIndex)
        {
            return doc.FirstSection.Body.Paragraphs[paraIndex];
        }
    }
}
