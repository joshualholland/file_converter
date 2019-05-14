using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;
using GroupDocs.Conversion;
using GroupDocs.Conversion.Config;
using GroupDocs.Conversion.Handler;

namespace PDF_DLL
{
    public class ConvertToPdf
    {
        public void Convert(string[] files)
        {
            // List of image files
            List<string> signatureList = new List<string>();
            // get logo file
            List<string> logo = new List<string>();
            for (int idx = 0; idx < files.Length; idx++)
            {
                if (Path.GetExtension(files[idx]) == ".doc" || Path.GetExtension(files[idx]) == ".docx")
                {
                    ConvertDoc(files[idx]);
                } else if (files[idx].Contains("compiled") || files[idx].Contains("reviewed"))
                {
                    signatureList.Add(files[idx]);
                }
                else if (Path.GetExtension(files[idx]) == ".xls" || Path.GetExtension(files[idx]) == ".xlsx")
                {
                    ConvertXls(files[idx]);
                }
                else if (files[idx].Contains("logo"))
                {
                    logo.Add(files[idx]);
                }
            }
            ConvertImg(signatureList);
            MergePdfs();
            AddLogo(logo);
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
                // Set font and size for title
                XFont title = new XFont("Open Sans", 24, XFontStyle.Regular);
                // Blue font color
                XBrush aldenBlue = new XSolidBrush(XColor.FromArgb(0, 82, 136));
                // muted font color
                XBrush muted = new XSolidBrush(XColor.FromArgb(153, 153, 153));
                // Set font and size
                XFont font = new XFont("Open Sans", 14, XFontStyle.Regular);
                // Put Signatures title on page
                info.DrawString("Signatures", title, XBrushes.Black, new XRect(0, 20, page.Width.Point, 0), XStringFormats.TopCenter);
                // Put compiled by signature form on page
                info.DrawString("Conversation #  ", font, aldenBlue, new XRect(96, 100, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString($"Compiled by: John Smith on {DateTime.Today.ToString("dd-MM-yyyy")} ", font, muted, new XRect(96, 125, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString("Signed:  ", font, muted, new XRect(96, 150, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                // Put Reviewed by signature form on page
                info.DrawString("Conversation #  ", font, aldenBlue, new XRect(96, 235, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString($"Reviewed by: Ted Cooper on {DateTime.Today.ToString("dd-MM-yyyy")} ", font, muted, new XRect(96, 260, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                info.DrawString("Signed:  ", font, muted, new XRect(96, 285, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                foreach (string image in images)
                {
                    if (image.Contains("compiled"))
                    {
                        // png for compiled by signature
                        DrawImage(info, image, 150, 150, 40, 40);
                    }
                    else if (image.Contains("reviewed"))
                    {
                        // png for reviewed by signature
                        DrawImage(info, image, 150, 285, 40, 40);
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
            string changeName = Path.GetFileNameWithoutExtension(filename) + "excel";
            convertedDocumentPath.Save(Path.ChangeExtension(changeName, ".pdf"));
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
                PdfDocument individualPdfs = PdfReader.Open(pdf, PdfDocumentOpenMode.Import);
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

        static void AddLogo(List<string> logoPath)
        {
            // Sets hardcoded directory for now
            string sourceDir = "../../doc/";
            // Gets combined pdf
            string[] files = Directory.GetFiles(sourceDir);
            List<string> combinedPdf = new List<string>();
            for (int dirinx = 0; dirinx < files.Length; dirinx++)
            {
                if (files[dirinx].Contains("combined"))
                {
                    combinedPdf.Add(files[dirinx]);
                }
            }
            // Open document to import pages
            PdfDocument individualPdfs = PdfReader.Open(combinedPdf[0], PdfDocumentOpenMode.Import);
            // Creates new Pdf
            PdfDocument pdfWithLogo = new PdfDocument();
            // font for footer
            XFont footer = new XFont("Open Sans", 9, XFontStyle.Regular);
            XBrush aldenBlue = new XSolidBrush(XColor.FromArgb(0, 82, 136));
            // Iterate through pages
            int count = individualPdfs.PageCount;
            for (int i = 0; i < count; i++)
            {
                PdfPage page = pdfWithLogo.AddPage(individualPdfs.Pages[i]);
                if (individualPdfs.Pages[i].Width <= 700)
                {
                    XGraphics gfx = XGraphics.FromPdfPage(page);
                    gfx.DrawString("Created by: ", footer, XBrushes.Black, new XRect(96, 770, page.Width.Point, 0), XStringFormats.BottomLeft);
                    gfx.DrawString("www.aldenone.com", footer, aldenBlue, new XRect(360, 770, page.Width.Point, 0), XStringFormats.BaseLineLeft);
                    DrawImage(gfx, logoPath[0], 150, 755, 67, 20);
                } else
                {
                    Console.WriteLine("not the right size");
                    Console.ReadKey();
                }
                
            }
            // Save final pdf
            pdfWithLogo.Save(sourceDir + "completedPdf.pdf");
        }
    }
}