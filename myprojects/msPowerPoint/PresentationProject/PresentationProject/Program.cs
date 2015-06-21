using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing; 
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Presentation;
//using Microsoft.Office.Interop.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Office2010.Drawing; 

namespace PresentationProject
{
    class Program
    {
        protected static System.Int64 cY;
        protected static int Count;

        protected static string headerText;
        protected static string contentText;

        static void Main(string[] args)
        {
            cY = 1353000L; // Set the Y coord for the first pic 
            Count = 0;

            headerText = "Test";
            contentText = "Some content.";


            string sourceFile = @"D:\Projects\word-doc\PresentationProject\Presentation1.pptx";
            string destFile = @"D:\Projects\word-doc\PresentationProject\Presentation1-used.pptx";
            System.IO.File.Copy(sourceFile, destFile, true);


            GetAllTextInSlide(destFile, 0);

            string anotherPic = @"D:\Projects\word-doc\image.jpg";
            InsertAPicture(destFile, 0, anotherPic);

            string anotherPic2 = @"D:\Projects\word-doc\image-old.jpg";
            InsertAPicture(destFile, 0, anotherPic2);

            string anotherPic3 = @"D:\Projects\word-doc\image2.jpg";
            InsertAPicture(destFile, 0, anotherPic3);

        }


        public static void GetAllTextInSlide(string presentationFile, int slideIndex)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, true))
            {
                GetAllTextInSlide(presentationDocument, slideIndex);
            }
        }

        public static void GetAllTextInSlide(PresentationDocument presentationDocument, int slideIndex)
        {
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            Presentation presentation = presentationPart.Presentation;

            if (presentation.SlideIdList != null)
            {
                DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                if (slideIndex < slideIds.Count)
                {
                    string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;
                    SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                    GetAllTextInSlide(slidePart, headerText, contentText);
                }

            }

            // return null;

        }

        public static void GetAllTextInSlide(SlidePart slidePart, string header, string content)
        {

            if (slidePart.Slide != null)
            {
                /*Header*/
                DocumentFormat.OpenXml.Drawing.Paragraph paragraph = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Paragraph>().First();

                DocumentFormat.OpenXml.Drawing.Text text = paragraph.Descendants<DocumentFormat.OpenXml.Drawing.Text>().First();
                text.Text = header;


                DocumentFormat.OpenXml.Drawing.TableCell cell = slidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TableCell>().First();

                /*Content*/
                DocumentFormat.OpenXml.Drawing.Text t = cell.Descendants<DocumentFormat.OpenXml.Drawing.Text>().First();
                t.Text = content;
            }
        }

        /*Image*/
        public static void InsertAPicture(string document, int slideIndex, string fileName)
        {
            using (PresentationDocument presentationDocument = PresentationDocument.Open(document, true))
            {
                PresentationPart presentationPart = presentationDocument.PresentationPart;

                Presentation presentation = presentationPart.Presentation;

                if (presentation.SlideIdList != null)
                {
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    if (slideIndex < slideIds.Count)
                    {
                        string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;                    

                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);

                        Slide slide = slidePart.Slide;

                        InsertImageInLastSlide(slidePart, slide, fileName);
                        
                    }

                }
            }
        }



        public static void InsertImageInLastSlide(SlidePart slidePart, Slide slide, string imagePath)
        {
            // Creates a Picture instance and adds its children. 
            P.Picture picture = new P.Picture();
            string embedId = string.Empty;
            embedId = "rId" + (slide.Elements<P.Picture>().Count() + 915 + Count).ToString();
            P.NonVisualPictureProperties nonVisualPictureProperties = new P.NonVisualPictureProperties(
                new P.NonVisualDrawingProperties() { Id = (UInt32Value)4U, Name = "Picture 5" },
                new P.NonVisualPictureDrawingProperties(new A.PictureLocks() { NoChangeAspect = true }),
                new ApplicationNonVisualDrawingProperties());

            P.BlipFill blipFill = new P.BlipFill();
            Blip blip = new Blip() { Embed = embedId };

            // Creates a BlipExtensionList instance and adds its children 
            BlipExtensionList blipExtensionList = new BlipExtensionList();
            BlipExtension blipExtension = new BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            UseLocalDpi useLocalDpi = new UseLocalDpi() { Val = false };
            useLocalDpi.AddNamespaceDeclaration("a14",
                "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension.Append(useLocalDpi);
            blipExtensionList.Append(blipExtension);
            blip.Append(blipExtensionList);

            Stretch stretch = new Stretch();
            FillRectangle fillRectangle = new FillRectangle();
            stretch.Append(fillRectangle);

            blipFill.Append(blip);
            blipFill.Append(stretch);

            // Creates a ShapeProperties instance and adds its children. 
            P.ShapeProperties shapeProperties = new P.ShapeProperties();

            A.Transform2D transform2D = new A.Transform2D();
            A.Offset offset = new A.Offset() { X = 9144000L, Y = cY };

            A.Extents extents = new A.Extents() { Cx = 2057400L, Cy = 1257300L };

            transform2D.Append(offset);
            transform2D.Append(extents);

            A.PresetGeometry presetGeometry = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList = new A.AdjustValueList();

            presetGeometry.Append(adjustValueList);

            shapeProperties.Append(transform2D);
            shapeProperties.Append(presetGeometry);

            picture.Append(nonVisualPictureProperties);
            picture.Append(blipFill);
            picture.Append(shapeProperties);

            slide.CommonSlideData.ShapeTree.AppendChild(picture);

            ImagePart imagePart = slidePart.AddImagePart(ImagePartType.Jpeg, embedId);
            using (FileStream stream = new FileStream(imagePath, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }

            cY += 1571625;
            Count++;
        }
    }
}
