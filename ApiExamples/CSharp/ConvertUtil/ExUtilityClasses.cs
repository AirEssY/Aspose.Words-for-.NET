﻿// Copyright (c) 2001-2014 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using Aspose.Words;
using NUnit.Framework;


namespace ApiExamples.ConvertUtil
{
    [TestFixture]
    public class ExUtilityClasses : ApiExampleBase
    {
        [Test]
        public void UtilityClassesUseControlCharacters()
        {
            string text = "test\r";
            //ExStart
            //ExFor:ControlChar
            //ExFor:ControlChar.Cr
            //ExFor:ControlChar.CrLf
            //ExId:UtilityClassesUseControlCharacters
            //ExSummary:Shows how to use control characters.
            // Replace "\r" control character with "\r\n"
            text = text.Replace(ControlChar.Cr, ControlChar.CrLf);
            //ExEnd
        }

        [Test]
        public void UtilityClassesConvertBetweenMeasurementUnits()
        {
            //ExStart
            //ExFor:ConvertUtil
            //ExId:UtilityClassesConvertBetweenMeasurementUnits
            //ExSummary:Shows how to specify page properties in inches.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            pageSetup.BottomMargin = Aspose.Words.ConvertUtil.InchToPoint(1.0);
            pageSetup.LeftMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            pageSetup.RightMargin = Aspose.Words.ConvertUtil.InchToPoint(1.5);
            pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);
            pageSetup.FooterDistance = Aspose.Words.ConvertUtil.InchToPoint(0.2);
            //ExEnd
        }

        [Test]
        public void MillimeterToPointEx()
        {
            //ExStart
            //ExFor:ConvertUtil.MillimeterToPoint
            //ExSummary:Shows how to specify page properties in millimeters.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.BottomMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(25.0);
            pageSetup.LeftMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.RightMargin = Aspose.Words.ConvertUtil.MillimeterToPoint(37.5);
            pageSetup.HeaderDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0);
            pageSetup.FooterDistance = Aspose.Words.ConvertUtil.MillimeterToPoint(5.0);

            builder.Writeln("Hello world.");
            builder.Document.Save(MyDir + "PageSetup.PageMargins Out.doc");
            //ExEnd
        }

        [Test]
        public void PointToInchEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PointToInch
            //ExSummary:Shows how to convert points to inches.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.InchToPoint(2.0);

            Console.WriteLine("The size of my top margin is {0} points, or {1} inches.",
                pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToInch(pageSetup.TopMargin));
            //ExEnd
        }

        [Test]
        public void PixelToPointEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToPoint(double)
            //ExFor:ConvertUtil.PixelToPoint(double, double)
            //ExSummary:Shows how to specify page properties in pixels with default and custom resolution.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetupNoDpi = builder.PageSetup;
            pageSetupNoDpi.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0);
            pageSetupNoDpi.BottomMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0);
            pageSetupNoDpi.LeftMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0);
            pageSetupNoDpi.RightMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0);
            pageSetupNoDpi.HeaderDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0);
            pageSetupNoDpi.FooterDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0);

            builder.Writeln("Hello world.");
            builder.Document.Save(MyDir + "PageSetup.PageMargins.DefaultResolution Out.doc");

            double myDpi = 150.0;

            Aspose.Words.PageSetup pageSetupWithDpi = builder.PageSetup;
            pageSetupWithDpi.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0, myDpi);
            pageSetupWithDpi.BottomMargin = Aspose.Words.ConvertUtil.PixelToPoint(100.0, myDpi);
            pageSetupWithDpi.LeftMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0, myDpi);
            pageSetupWithDpi.RightMargin = Aspose.Words.ConvertUtil.PixelToPoint(150.0, myDpi);
            pageSetupWithDpi.HeaderDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0, myDpi);
            pageSetupWithDpi.FooterDistance = Aspose.Words.ConvertUtil.PixelToPoint(20.0, myDpi);

            builder.Document.Save(MyDir + "PageSetup.PageMargins.CustomResolution Out.doc");
            //ExEnd
        }

        [Test]
        public void PointToPixelEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PointToPixel(double)
            //ExFor:ConvertUtil.PointToPixel(double, double)
            //ExSummary:Shows how to use convert points to pixels with default and custom resolution.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = Aspose.Words.ConvertUtil.PixelToPoint(2.0);

            double myDpi = 192.0;

            Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with default resolution.",
                pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToPixel(pageSetup.TopMargin));

            Console.WriteLine("The size of my top margin is {0} points, or {1} pixels with custom resolution.",
                pageSetup.TopMargin, Aspose.Words.ConvertUtil.PointToPixel(pageSetup.TopMargin, myDpi));
            //ExEnd
        }

        [Test]
        public void PixelToNewDpiEx()
        {
            //ExStart
            //ExFor:ConvertUtil.PixelToNewDpi
            //ExSummary:Shows how to check how an amount of pixels changes when the dpi is changed.
            Aspose.Words.Document doc = new Aspose.Words.Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            Aspose.Words.PageSetup pageSetup = builder.PageSetup;
            pageSetup.TopMargin = 72;
            double oldDpi = 92.0;
            double newDpi = 192.0;

            Console.WriteLine("{0} pixels at {1} dpi becomes {2} pixels at {3} dpi.",
                pageSetup.TopMargin, oldDpi, Aspose.Words.ConvertUtil.PixelToNewDpi(pageSetup.TopMargin, oldDpi, newDpi), newDpi);
            //ExEnd
        }
    }
}
