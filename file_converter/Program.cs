using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using System.Web;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using GroupDocs.Conversion;
using GroupDocs.Conversion.Config;
using GroupDocs.Conversion.Handler;


namespace file_converter
{
    class Program
    {
        static void Main(string[] args)
        {
            // hard coded directory
            string dir = "../../doc/";
            // List of image files
            List<string> imgList = new List<string>();
            // List of all files
            string[] files = Directory.GetFiles(dir);
            for (int idx = 0; idx < files.Length; idx++)
            {
                if (Path.GetExtension(files[idx]) == ".doc" || Path.GetExtension(files[idx]) == ".docx")
                {
                    ConvertDoc(files[idx]);
                }
                else if (Path.GetExtension(files[idx]) == ".jpg" || Path.GetExtension(files[idx]) == ".png" || Path.GetExtension(files[idx]) == ".jpeg")
                {
                    imgList.Add(files[idx]);
                    // ConvertImg(files[idx]);
                } else if (Path.GetExtension(files[idx]) == ".xls" || Path.GetExtension(files[idx]) == ".xlsx")
                {
                    ConvertXls(files[idx]);
                }
                else
                {
                    Console.WriteLine(files[idx]);
                    Console.ReadKey();
                }
            }

            ConvertImg(imgList);
       
            MergePdfs();
        }

        static void ConvertDoc(string filename)
        {
            // Where to receive and save file
            var conversionConfig = new ConversionConfig { StoragePath = filename, OutputPath = Path.GetDirectoryName(filename) };
            var conversionHandler = new ConversionHandler(conversionConfig);
            // Saves new file as pdf
            var saveOptions = new GroupDocs.Conversion.Options.Save.PdfSaveOptions();
            // Converts existing doc to pdf
            var convertedDocumentPath = conversionHandler.Convert(filename, saveOptions);
            // change name so that xls and doc files of the same name can both be saved
            string changeName = Path.GetFileNameWithoutExtension(filename) + "word";
            // save new pdf with ".pdf" ext
            convertedDocumentPath.Save(Path.ChangeExtension(changeName, ".pdf"));

        }

        static void ConvertImg(List<string> images)
        {
            {
                // Create new pdf document
                PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();
                // Add image to page of pdf
                PdfPage page = document.AddPage();
                // Put xobject on page (signature info)
                XGraphics info = XGraphics.FromPdfPage(page);
                // Set font and size
                XFont font = new XFont("Arial", 14, XFontStyle.Regular);
                // Put Signatures title on page
                info.DrawString("Signatures", font, XBrushes.Black, new XRect(0, 20, page.Width.Point, 0), XStringFormats.TopCenter);
                // Put compiled by signature form on page
                info.DrawString("Conversation #  ", font, XBrushes.Black, new XRect(0, 80, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString($"Compiled by: John Smith on {DateTime.Today.ToString("dd-MM-yyyy")} ", font, XBrushes.Black, new XRect(0, 105, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString("Signed:  ", font, XBrushes.Black, new XRect(0, 125, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                // Put Reviewed by signature form on page
                info.DrawString("Conversation #  ", font, XBrushes.Black, new XRect(0, 185, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString($"Reviewed by: Ted Cooper on {DateTime.Today.ToString("dd-MM-yyyy")} ", font, XBrushes.Black, new XRect(0, 210, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString("Signed:  ", font, XBrushes.Black, new XRect(0, 235, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                foreach (string image in images)
                {
                    if (image.Contains("compiled"))
                    {
                        // png for compiled by signature
                        DrawImage(info, image, 60, 115, 40, 40);
                    }
                    else if (image.Contains("reviewed"))
                    {
                        // png for reviewed by signature
                        DrawImage(info, image, 60, 225, 40, 40);
                    }
                }
                

                // Save pdf image in indiv folder
                document.Save("../../doc/z_signature.pdf");
            }
            
            
        }

        private static void DrawImage(XGraphics gfx, string jpegPath, int x, int y, int width, int height)
        {
            // Put xobject on page 
            XImage image = XImage.FromFile(jpegPath);
            gfx.DrawImage(image, x, y, width, height);

        }

        static void ConvertXls(string filename)
        {
            var conversionConfig = new ConversionConfig { StoragePath = filename, OutputPath = Path.GetDirectoryName(filename) };
            var conversionHandler = new ConversionHandler(conversionConfig);
            var saveOptions = new GroupDocs.Conversion.Options.Save.PdfSaveOptions();
            var convertedDocumentPath = conversionHandler.Convert(filename, saveOptions);

            convertedDocumentPath.Save(Path.GetFileNameWithoutExtension(filename) + ".pdf");
        }

        static void MergePdfs()
        {
            // Sets hardcoded directory for now
            string sourceDir = "../../doc/";
            // List for pdfs
            List<string> pdfs = new List<string>();
            // Gets files
            string[] files = Directory.GetFiles(sourceDir);
            for (int dirinx = 0; dirinx < files.Length; dirinx++)
            {
                if (Path.GetExtension(files[dirinx]) == ".pdf")
                {
                    pdfs.Add(files[dirinx]);
                } 
            }
            // Creates new Pdf
            PdfSharp.Pdf.PdfDocument combinedPdf = new PdfSharp.Pdf.PdfDocument();
            // Iterates through files
            foreach (string pdf in pdfs)
            {
                // Open document to import pages
                PdfSharp.Pdf.PdfDocument individualPdfs = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);

                // Iterate through pages
                int count = individualPdfs.PageCount;
                for (int i = 0; i < count; i++)
                {
                    PdfPage page = individualPdfs.Pages[i];
                    combinedPdf.AddPage(page);
                }
            }

            // Save document
            const string filename = "combinedPdfTest1.pdf";
            combinedPdf.Save(sourceDir + filename);
        }
        
    }
}
