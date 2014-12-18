using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Saving;
using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using Aspose.Pdf.Devices;
using System.IO;
using Aspose.Diagram;
using Aspose.Slides;
using Aspose.Cells;
using Aspose.Cells.Rendering;
using System.Net.Mail;
using Aspose.Email.Mail;

namespace Aspose.POC
{
    public class Program
    {
        public static void Main(string[] args)
        {
            MailToImage();
            //ExcelToImage();
            //PowerPointSlideToImage();
            //VisioToImage();
            //DocumentToImage();
            //PdfToImage();
        }

        private static void VisioToImage()
        {
            Aspose.Diagram.Diagram diagram = new Aspose.Diagram.Diagram(@"C:\test.vsd");

            Aspose.Diagram.Saving.ImageSaveOptions options = new Aspose.Diagram.Saving.ImageSaveOptions(SaveFileFormat.JPEG);

            diagram.Save(@"C:\VISIO.jpg", options);

        }

        private static void PdfToImage()
        {
            Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(@"C:\test.pdf");

            for (int pageCount = 1; pageCount <= pdfDocument.Pages.Count; pageCount++)
            {
                using (FileStream imageStream = new FileStream(@"C:\PDFimage" + pageCount + ".jpg", FileMode.Create))
                {
                    // Create Resolution object
                    Resolution resolution = new Resolution(300);
                    // Create JPEG device with specified attributes (Width, Height, Resolution, Quality)
                    // where Quality [0-100], 100 is Maximum
                    JpegDevice jpegDevice = new JpegDevice(resolution, 100);

                    // Convert a particular page and save the image to stream
                    jpegDevice.Process(pdfDocument.Pages[pageCount], imageStream);
                    // Close stream
                    imageStream.Close();
                }
            }
            //Convert each page to PDF file

        }

        private static void DocumentToImage()
        {

            Document doc = new Document(@"C:\test.doc");
            var inputFileName = @"C:\test.docx";

            Aspose.Cloud.Words.DocumentBuilder builder = new Aspose.Cloud.Words.DocumentBuilder();

            builder.insertWatermarkImage(fileName, watermarkImage, rotationAngle);

            Aspose.Words.Saving.ImageSaveOptions options = new Aspose.Words.Saving.ImageSaveOptions(Aspose.Words.SaveFormat.Jpeg);
            options.PageCount = 1;
            for (int pageIndex = 0; pageIndex < doc.PageCount; pageIndex++)
            {
                string outputFileName = @"C:\" + string.Format("DOCUMENT{0}_{1}.jpeg", "Test", pageIndex + 1);
                options.PageIndex = pageIndex;
                
                doc.Save(outputFileName, options);
            }

        }

        private static void PowerPointSlideToImage()
        {
            using (Aspose.Slides.Presentation pres = new Aspose.Slides.Presentation(@"C:\test.pptx"))
            {
                for (int i = 0; i < pres.Slides.Count; i++)
                {
                    var sld = pres.Slides[i];
                    int desiredX = 1200;
                    int desiredY = 800;

                    float ScaleX = (float)(1.0 / pres.SlideSize.Size.Width) * desiredX;
                    float ScaleY = (float)(1.0 / pres.SlideSize.Size.Height) * desiredY;

                    var bmp = sld.GetThumbnail(ScaleX, ScaleY);

                    bmp.Save(@"C:\pptSlide" + i + ".jpg", System.Drawing.Imaging.ImageFormat.Jpeg);
                }
            }

        }

        private static void ExcelToImage()
        {

            Workbook book = new Workbook(@"C:\test.xls");

            //for (int i = 0; i < book.Worksheets.Count; i++) //todo loopshrough worksheets
            //{
            Worksheet sheet = book.Worksheets[0];

            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions();
            imgOptions.ImageFormat = System.Drawing.Imaging.ImageFormat.Jpeg;
            imgOptions.OnePagePerSheet = true;
            SheetRender sr = new SheetRender(sheet, imgOptions);
            Bitmap bitmap = sr.ToImage(0);
            bitmap.Save(@"C:\excelworksheet" + 1 + ".jpg");
            //}
        }

        private static void MailToImage()
        {

            Aspose.Email.Mail.MailMessage msg = Aspose.Email.Mail.MailMessage.Load(@"C:\test.msg", MessageFormat.Msg);

            MemoryStream msgStream = new MemoryStream();
            msg.Save(msgStream, MailMessageSaveType.MHtmlFromat);
            msgStream.Position = 0;

            Document msgDocument = new Document(msgStream);

            msgDocument.Save(@"C:\Outlook-Aspose.jpeg", Aspose.Words.SaveFormat.Jpeg);

        }

    }
}

