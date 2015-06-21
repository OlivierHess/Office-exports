using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;



namespace WordProject
{
    class Program
    {

        static void Main(string[] args)
        {
            string header = "HeaderTest";
            
            string content = "Some content.";

            string sourceFile = @"D:\Projects\word-doc\Doc3.docx";
            string destFile = @"D:\Projects\word-doc\Doc-used.docx";
            System.IO.File.Copy(sourceFile, destFile, true);

            ChangeHeader(destFile, header);

            WriteContent(destFile, content);

            string fileName = @"D:\Projects\word-doc\image-old.jpg";
            InsertAPicture(destFile, fileName);

            string anotherPic = @"D:\Projects\word-doc\image.jpg";
            InsertAPicture(destFile, anotherPic);

            string fileName2 = @"D:\Projects\word-doc\image-old2.jpg";
            InsertAPicture(destFile, fileName2);
            
         }

        /*Image*/
        public static void InsertAPicture(string document, string fileName)
        {
            using (WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(document, true))
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
            var element =
        new Drawing(
            new DW.Inline(
                new DW.Extent() { Cx = 990000L, Cy = 792000L },
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
                                    CompressionState =
                                    A.BlipCompressionValues.Print
                                },
                                new A.Stretch(
                                    new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset() { X = 0L, Y = 0L },
                                    new A.Extents() { Cx = 990000L, Cy = 792000L }),
                                new A.PresetGeometry(
                                    new A.AdjustValueList()
                                ) { Preset = A.ShapeTypeValues.Rectangle }))
                    ) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
            )
            {
                DistanceFromTop = (UInt32Value)0U,
                DistanceFromBottom = (UInt32Value)0U,
                DistanceFromLeft = (UInt32Value)0U,
                DistanceFromRight = (UInt32Value)0U,
                EditId = "50D07946"
            });

            //Insert to the 2nd column
            Table table = wordDoc.MainDocumentPart.Document.Body.Elements<Table>().First();
            TableRow row = table.Elements<TableRow>().ElementAt(0);
            TableCell cell = row.Elements<TableCell>().ElementAt(1);

            Paragraph p = cell.AppendChild(new Paragraph());
            Run r = p.AppendChild(new Run());
            r.AppendChild(element);

            //Add an extra line break
            cell.AppendChild(new Paragraph());
        }

        public static void WriteContent(string filepath, string txt)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
            {
                Table table = doc.MainDocumentPart.Document.Body.Elements<Table>().First();
                TableRow row = table.Elements<TableRow>().First();
                TableCell cell = row.Elements<TableCell>().First();
                SdtBlock block = cell.Elements<SdtBlock>().First();
                SdtContentBlock container = block.Elements<SdtContentBlock>().First();
                Paragraph p = container.Elements<Paragraph>().First();
                Run r = p.Elements<Run>().First();
                Text t = r.Elements<Text>().First();
                t.Text = txt;
            }
        }

        public static void ChangeHeader(string filepath, string txt)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
            {
                Paragraph p = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().First();
                Run r = p.Elements<Run>().First();
                Text t = r.Elements<Text>().First();
                t.Text = txt;
            }
        }


    }
}
