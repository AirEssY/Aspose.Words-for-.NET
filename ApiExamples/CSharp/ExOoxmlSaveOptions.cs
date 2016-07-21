// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using NUnit.Framework;

namespace ApiExamples
{
    using System;
    using System.IO;

    using Aspose.Words;
    using Aspose.Words.Drawing;
    using Aspose.Words.Saving;

    [TestFixture]
    internal class ExOoxmlSaveOptions : ApiExampleBase
    {
        [Test]
        public void Iso29500Strict()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            doc.CompatibilityOptions.OptimizeFor(MsWordVersion.Word2003);
            
            Shape image = builder.InsertImage(MyDir + @"dotnet-logo.png");

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Vml, shape.MarkupLanguage);
            }

            OoxmlSaveOptions saveOptions = new OoxmlSaveOptions();
            saveOptions.Compliance = OoxmlCompliance.Iso29500_2008_Strict;
            saveOptions.SaveFormat = SaveFormat.Docx;

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, saveOptions);

            // Loop through all single shapes inside document.
            foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true))
            {
                Assert.AreEqual(ShapeMarkupLanguage.Dml, shape.MarkupLanguage);
            }
        }
    }
}
