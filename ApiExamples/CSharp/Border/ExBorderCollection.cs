﻿// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;


namespace ApiExamples.Border
{
    [TestFixture]
    public class ExBorderCollection : ApiExampleBase
    {
        [Test]
        public void GetEnumeratorEx()
        {
            //ExStart
            //ExFor:BorderCollection.GetEnumerator
            //ExSummary:Shows how to enumerate all borders in a collection.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            var enumerator = borders.GetEnumerator();
            while (enumerator.MoveNext())
            {
                // Do something useful.
                Aspose.Words.Border b = (Aspose.Words.Border)enumerator.Current;
                b.Color = System.Drawing.Color.RoyalBlue;
                b.LineStyle = LineStyle.Double;
            }

            doc.Save(MyDir + "Document.ChangedColourBorder.doc");
            //ExEnd
        }

        [Test]
        public void ClearFormattingEx()
        {
            //ExStart
            //ExFor:BorderCollection.ClearFormatting
            //ExSummary:Shows how to remove all borders from a paragraph at once.
            Aspose.Words.Document doc = new Aspose.Words.Document(MyDir + "Document.Borders.doc");
            DocumentBuilder builder = new DocumentBuilder(doc);
            BorderCollection borders = builder.ParagraphFormat.Borders;

            borders.ClearFormatting();
            //ExEnd
        }
    }
}