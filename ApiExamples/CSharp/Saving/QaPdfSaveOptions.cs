﻿using Aspose.Words;
using Aspose.Words.Saving;
using NUnit.Framework;

namespace ApiExamples.Saving
{
    using Document = Aspose.Words.Document;

    [TestFixture]
    internal class QaPdfSaveOptions : ApiExampleBase
    {
        //Note: Test doesn't containt validation result, because it's difficult 
        //For validation result, you can save the document to pdf file and check out, that all bookmarks are created correctly for missing headings
        [Test]
        public void CreateMissingOutlineLevels()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();

            //Set maximum value of levels of headings
            pdfSaveOptions.OutlineOptions.HeadingsOutlineLevels = 9;
            pdfSaveOptions.OutlineOptions.CreateMissingOutlineLevels = true;
            pdfSaveOptions.OutlineOptions.ExpandedOutlineLevels = 9;

            pdfSaveOptions.SaveFormat = SaveFormat.Pdf;

            doc.Save(MyDir + "CreateMissingOutlineLevels_OUT.pdf", pdfSaveOptions);
        }

        //Note: Test doesn't containt validation result, because it's difficult
        //For validation result, you can add some shapes to the document and assert, that the DML shapes are render correctly
        [Test]
        public void DrawingMl()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            PdfSaveOptions pdfSaveOptions = new PdfSaveOptions();
            pdfSaveOptions.DmlRenderingMode = DmlRenderingMode.DrawingML;

            doc.Save(MyDir + "DrawingMl_OUT.pdf", pdfSaveOptions);
        }
    }
}
