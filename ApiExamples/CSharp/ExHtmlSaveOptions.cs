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

    using Document = Aspose.Words.Document;
    using HtmlSaveOptions = Aspose.Words.Saving.HtmlSaveOptions;
    using SaveFormat = Aspose.Words.SaveFormat;

    [TestFixture]
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        //For assert this test you need to open html docs and they shouldn't have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions
            {
                SaveFormat = saveFormat, 
                ExportPageMargins = true
            };

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    doc.Save(MyDir + "ExportPageMargins.html", htmlSaveOptions);
                    break;
                case SaveFormat.Mhtml:
                    doc.Save(MyDir + "ExportPageMargins.Mhtml", htmlSaveOptions);
                    break;
                case SaveFormat.Epub:
                    doc.Save(MyDir + "ExportPageMargins.Epub", htmlSaveOptions); //There is draw images bug with epub. Need write to NSezganov
                    break;
            }
        }

        [Test]
        [TestCase(true)]
        [TestCase(false)]
        public void ExportUrlForLinkedImage(bool export)
        {
            Document doc = new Document(MyDir + "ExportUrlForLinkedImage.docx");
            
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            saveOptions.ExportOriginalUrlForLinkedImages = export;

            doc.Save(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", saveOptions);

            String[] dirFiles = Directory.GetFiles(MyDir + @"\Artifacts\", "ExportUrlForLinkedImage.001.png", SearchOption.AllDirectories);

            if (dirFiles.Length == 0)
            {
                this.FindTextInFile(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", "<img src=\"http://www.aspose.com/images/aspose-logo.gif\"");
            }
            else
            {
                this.FindTextInFile(MyDir + @"\Artifacts\ExportUrlForLinkedImage.html", "<img src=\"ExportUrlForLinkedImage.001.png\"");
            }
        }

        //ToDo: Change location to helper
        private void FindTextInFile(string path, string expression)
        {
            using (var sr = new StreamReader(path))
            {
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine();

                    if (String.IsNullOrEmpty(line)) continue;

                    if (line.Contains(expression))
                    {
                        Console.WriteLine(line);
                        Assert.Pass();
                    }
                    else
                    {
                        Assert.Fail();
                    }
                }
            }
        }
    }
}
