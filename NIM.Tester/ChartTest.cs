using DocumentFormat.OpenXml.Office.Drawing;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using asp = Aspose.Cells;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using System.Windows.Media.Imaging;
using NIM.Utilty;
using  w = DocumentFormat.OpenXml.Wordprocessing;

namespace ExcelTestor
{
    [TestClass]
    public class ChartTest
    {

        [TestMethod]
        public void ChartToImageTest()
        {



            var docName = @"O:\nim\images\2017.4.24.xlsx";
            var workbook = new asp.Workbook(docName);
            //workbook.CalculateFormula();

            ChartHelper.ToImage(workbook.Worksheets[1].Charts[0], @"O:\nim\images\david.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            //workbook.Worksheets[1].Charts[0].ToImage(@"O:\cic\images\david.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            //InsertAPicture(@"O:\cic\images\David charts.docx", @"O:\cic\images\david.jpg");

        }
        public static void InsertAPicture(string document, string fileName)
        {
            using (WordprocessingDocument wordprocessingDocument =
                WordprocessingDocument.Open(document, true))
            {
                MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;

                ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Jpeg);

                using (FileStream stream = new FileStream(fileName, FileMode.Open))
                {
                    imagePart.FeedData(stream);
                }

                AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart));
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId)
        {
            //https://stackoverflow.com/questions/8082980/inserting-image-into-docx-using-openxml-and-setting-the-size

            var table = wordDoc.MainDocumentPart.Document.Body.Elements<Table>().First();
            var cell = table.Elements<TableRow>().First().Elements<TableCell>().First();


            var fileName = @"O:\cic\images\david.jpg";

            var img = new BitmapImage(new Uri(fileName, UriKind.RelativeOrAbsolute));
            var widthPx = img.PixelWidth;
            var heightPx = img.PixelHeight;
            var horzRezDpi = img.DpiX;
            var vertRezDpi = img.DpiY;
            const int emusPerInch = 914400;
            const int emusPerCm = 360000;
            var maxWidthCm = 16.51;
            var widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
            var heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);
            var maxWidthEmus = (long)(maxWidthCm * emusPerCm);
            if (widthEmus > maxWidthEmus)
            {
                var ratio = (heightEmus * 1.0m) / widthEmus;
                widthEmus = maxWidthEmus;
                heightEmus = (long)(widthEmus * ratio);
            }

            //widthEmus = 5382931L;
            //heightEmus = 3466036L;

            // Define the reference of the image.
            var element =
                 new DocumentFormat.OpenXml.Wordprocessing.Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = widthEmus, Cy = heightEmus },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                       "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState = A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = widthEmus, Cy = heightEmus }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });


            var garagraph = new Paragraph(new Run(element));
            cell.Append(garagraph);
            // Append the reference to body, the element should be in a Run.
            //  wordDoc.MainDocumentPart.Document.Body.AppendChild(new Paragraph(new Run(element)));
        }




    }
}
