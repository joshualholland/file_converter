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
                    ConvertImg(files[idx]);
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
       
            MergePdfs();
        }

        static void ConvertDoc(string filename)
        {
            var conversionConfig = new ConversionConfig { StoragePath = filename, OutputPath = Path.GetDirectoryName(filename) };
            var conversionHandler = new ConversionHandler(conversionConfig);
            var saveOptions = new GroupDocs.Conversion.Options.Save.PdfSaveOptions();
            var convertedDocumentPath = conversionHandler.Convert(filename, saveOptions);
            string changeName = Path.GetFileNameWithoutExtension(filename) + "word";
            convertedDocumentPath.Save(Path.ChangeExtension(changeName, ".pdf"));

        }

        static void ConvertImg(string filename)
        {
            // Create new pdf document
            PdfSharp.Pdf.PdfDocument document = new PdfSharp.Pdf.PdfDocument();
            // Add image to page of pdf
            PdfPage page = document.AddPage();
            // Get xobject from image
            XGraphics gfx = XGraphics.FromPdfPage(page);
            DrawImage(gfx, filename, 0, 0, 50, 50);
            // Save pdf image in indiv folder
            document.Save("../../doc/png.pdf");
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
