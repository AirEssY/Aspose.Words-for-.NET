﻿using CSharp.Loading_Saving;
using CSharp.Mail_Merge;
using CSharp.Programming_Documents.Find_and_Replace;
using CSharp.Programming_Documents.Joining_and_Appending;
using CSharp.Programming_Documents.Bookmarks;
using CSharp.Programming_Documents.Comments;
using CSharp.Programming_Documents.Working_With_Document;
using CSharp.Programming_Documents.Working_with_Fields;
using CSharp.Programming_Documents.Working_with_Ranges;
using CSharp.Programming_Documents.Working_with_Images;
using CSharp.Programming_Documents.Working_with_Styles;
using CSharp.Programming_Documents.Working_with_Tables;
using CSharp.Programming_Documents.Working_with_Sections;
using CSharp.Quick_Start;
using CSharp.Rendering_and_Printing;
using CSharp.LINQ;
using DocumentExplorerExample;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace CSharp
{
    class RunExamples
    {
        [STAThread]
        public static void Main()
        {
            Console.WriteLine("Open RunExamples.cs. In Main() method, Un-comment the example that you want to run");
            Console.WriteLine("=====================================================");
            // Un-comment the one you want to try out

            // =====================================================
            // =====================================================
            // Quick Start
            // =====================================================
            // =====================================================

            //AppendDocuments.Run();
            //ApplyLicense.Run();           
            //FindAndReplace.Run();
            //CSharp.Quick_Start.HelloWorld.Run();            
            //SimpleMailMerge.Run();
            //UpdateFields.Run();
            //WorkingWithNodes.Run();

            //// =====================================================
            //// =====================================================
            //// Loading and Saving
            //// =====================================================
            //// =====================================================

            //OpenEncryptedDocument.Run();
            //LoadAndSaveToDisk.Run();
            //LoadAndSaveToStream.Run();
            //CreateDocument.Run();
            //CheckFormat.Run();
            //SplitIntoHtmlPages.Run();
            //LoadTxt.Run();
            //PageSplitter.Run();
            //ImageToPdf.Run();
            //SpecifySaveOption.Run();
            //AccessAndVerifySignature.Run();
            //Doc2Pdf.Run();
            //DigitallySignedPdf.Run();
            //ConvertDocumentToByte.Run();
            //ConvertDocumentToEPUB.Run();
            //ConvertDocumentToHtmlWithRoundtrip.Run();

            //// =====================================================
            //// =====================================================
            //// Programming with Documents
            //// =====================================================
            //// =====================================================

            //// Joining and Appending
            //// =====================================================
            //SimpleAppendDocument.Run();
            //KeepSourceFormatting.Run();
            //UseDestinationStyles.Run();
            //JoinContinuous.Run();
            //JoinNewPage.Run();
            //RestartPageNumbering.Run();
            //LinkHeadersFooters.Run();
            //UnlinkHeadersFooters.Run();
            //RemoveSourceHeadersFooters.Run();
            //DifferentPageSetup.Run();
            //ConvertNumPageFields.Run();
            //ListUseDestinationStyles.Run();
            //ListKeepSourceFormatting.Run();
            //KeepSourceTogether.Run();
            //BaseDocument.Run();
            //UpdatePageLayout.Run();
            //AppendDocumentManually.Run();
            //PrependDocument.Run();

            //// Find and Replace
            //// =====================================================
            //FindAndHighlight.Run();
            //ReplaceTextWithField.Run();

            //// Bookmarks
            //// =====================================================
            //CopyBookmarkedText.Run();
            //UntangleRowBookmarks.Run();
            //BookmarkTable.Run();
            //BookmarkNameAndText.Run();
            //AccessBookmarks.Run();
            //CreateBookmark.Run();

            //// Comments
            //// =====================================================
            //ProcessComments.Run();

            //// Document
            //// =====================================================
            //ExtractContentBetweenParagraphs.Run();
            //ExtractContentBetweenBlockLevelNodes.Run();
            //ExtractContentBetweenParagraphStyles.Run();
            //ExtractContentBetweenRuns.Run();
            //ExtractContentUsingField.Run();
            //ExtractContentBetweenBookmark.Run();
            //ExtractContentBetweenCommentRange.Run();
            //PageNumbersOfNodes.Run();
            //RemoveBreaks.Run();
            //CloningDocument.Run();
            //ProtectDocument.Run();
            //AccessStyles.Run();
            //GetVariables.Run();
            //SetViewOption.Run();

            //// Fields
            //// =====================================================
            //InsertNestedFields.Run();
            //RemoveField.Run();
            //ConvertFieldsInDocument.Run();
            //ConvertFieldsInBody.Run();
            //ConvertFieldsInParagraph.Run();
            //ChangeLocale.Run();
            //UpdateDocFields.Run();
            //InsertField.Run();
            //InsertMergeFieldUsingDOM.Run();
            //InsertMailMergeAddressBlockFieldUsingDOM.Run();
            //InsertAdvanceFieldWithOutDocumentBuilder.Run();
            //InsertASKFieldWithOutDocumentBuilder.Run();
            //InsertAuthorField.Run();
            //InsertFormFields.Run();
            //FormFieldsGetFormFieldsCollection.Run();
            //FormFieldsGetByName.Run();
            //FormFieldsWorkWithProperties.Run();
            
            //// Images
            //// =====================================================
            //AddImageToEachPage.Run();
            //AddWatermark.Run();
            //CompressImages.Run();

            //// Ranges
            //// =====================================================
            //RangesGetText.Run();
            //RangesDeleteText.Run();
            

            //// Styles
            //// =====================================================
            //ExtractContentBasedOnStyles.Run();

            //// Tables
            //// =====================================================
            //AutoFitTableToWindow.Run();
            //AutoFitTableToContents.Run();
            //AutoFitTableToFixedColumnWidths.Run();
            //InsertTableUsingDocumentBuilder.Run();
            //InsertTableDirectly.Run();
            //CloneTable.Run();
            //InsertTableFromHtml.Run();
            //ApplyFormatting.Run();
            //SpecifyHeightAndWidth.Run();
            //ApplyStyle.Run();
            //ExtractText.Run();
            //FindingIndex.Run();
            //AddRemoveColumn.Run();
            //RepeatRowsOnSubsequentPages.Run();
            //JoiningAndSplittingTable.Run();            

            //// Sections
            //// =====================================================
            //SectionsAccessByIndex.Run();
            //AddDeleteSection.Run();
            //AppendSectionContent.Run();
            //DeleteSectionContent.Run();
            //DeleteHeaderFooterContent.Run();
            //CloneSection.Run();
            //CopySection.Run();
            

            //// =====================================================
            //// =====================================================
            //// MailMerge and Reporting
            //// =====================================================
            //// =====================================================

            //ApplyCustomLogicToEmptyRegions.Run();
            //LINQtoXMLMailMerge.Run();
            //MailMergeFormFields.Run();
            //MultipleDocsInMailMerge.Run();
            //NestedMailMerge.Run();
            //RemoveEmptyRegions.Run();
            //XMLMailMerge.Run();            

            //// =====================================================
            //// =====================================================
            //// Rendering and Printing
            //// =====================================================
            //// =====================================================

            //DocumentLayoutHelper.Run();
            //EnumerateLayoutElements.Run();
            //DocumentPreviewAndPrint.Run();
            //ImageColorFilters.Run();
            //RenderShape.Run();
            //SaveAsMultipageTiff.Run();
            //ReadActiveXControlProperties.Run();

            //// =====================================================
            //// =====================================================
            //// Viewers and Visualizers
            //// =====================================================
            //// =====================================================

            //MainForm.Run();

            //// =====================================================
            //// =====================================================
            //// LINQ
            //// =====================================================
            //// =====================================================

            //CSharp.LINQ.HelloWorld.Run();
            //SingleRow.Run();
            //InParagraphList.Run();
            //BulletedList.Run();
            //NumberedList.Run();
            //MulticoloredNumberedList.Run();
            //CommonList.Run();
            //InTableList.Run();
            //InTableAlternateContent.Run();
            //CommonMasterDetail.Run();
            //InTableMasterDetail.Run();
            //InTableWithFilteringGroupingSorting.Run();
            //PieChart.Run();
            //ScatterChart.Run();
            //BubbleChart.Run();
            //ChartWithFilteringGroupingOrdering.Run();
         
            // Stop before exiting
            Console.WriteLine("\n\nProgram Finished. Press any key to exit....");
            Console.ReadKey();
        }

        public static String GetDataDir_LINQ()
        {
            return Path.GetFullPath(GetDataDir_Data() + "LINQ/");
        }
        public static String GetDataDir_LoadingAndSaving()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Loading-and-Saving/");
        }

        public static String GetDataDir_JoiningAndAppending()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Joining-Appending/");
        }

        public static String GetDataDir_FindAndReplace()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Find-Replace/");
        }

        public static String GetDataDir_WorkingWithBookmarks()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Bookmarks/");
        }

        public static String GetDataDir_WorkingWithComments()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Comments/");
        }

        public static String GetDataDir_WorkingWithDocument()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Document/");
        }

        public static String GetDataDir_WorkingWithFields()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Fields/");
        }
        public static String GetDataDir_WorkingWithRanges()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Ranges/");
        }

        public static String GetDataDir_WorkingWithImages()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Images/");
        }

        public static String GetDataDir_WorkingWithStyles()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Styles/");
        }

        public static String GetDataDir_WorkingWithTables()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Tables/");
        }
        public static String GetDataDir_WorkingWithSections()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Programming-Documents/Sections/");
        }
        public static String GetDataDir_MailMergeAndReporting()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Mail-Merge/");
        }

        public static String GetDataDir_QuickStart()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Quick-Start/");
        }

        public static String GetDataDir_RenderingAndPrinting()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Rendering-Printing/");
        }

        public static String GetDataDir_ViewersAndVisualizers()
        {
            return Path.GetFullPath(GetDataDir_Data() + "Viewers-Visualizers/");
        }
        private static string GetDataDir_Data()
        {
            var parent = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
            string startDirectory = null;
            if (parent != null)
            {
                var directoryInfo = parent.Parent;
                if (directoryInfo != null)
                {
                    startDirectory = directoryInfo.FullName;
                }
            }
            else
            {
                startDirectory = parent.FullName;
            }
            return Path.Combine(startDirectory, "Data\\");
        }
        public static string GetOutputFilePath(String inputFilePath)
        {
            string extension = Path.GetExtension(inputFilePath);
            string filename = Path.GetFileNameWithoutExtension(inputFilePath);
            return filename + "_out_" + extension;
        }
    }
}
