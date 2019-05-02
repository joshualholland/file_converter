using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Aspose.Words;
using PdfSharp;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using PdfSharp.Pdf.IO;

namespace file_converter
{
    class Program
    {
        static void Main(string[] args)
        {
            docToPdf word = new docToPdf();
            word.Convert();
            imagesToPdf image = new imagesToPdf();
            image.Convert();

            MergePdfs();
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
            PdfDocument combinedPdf = new PdfDocument();
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
        
    }

    class docToPdf
    {
        public void Convert()
        {
            // Convert doc to pdf
            string path = "../../doc/";
            string fileName1 = path + "resume.docx";
            Document doc = new Document(fileName1);
            doc.Save(path + "DocumentToPdf.pdf", SaveFormat.Pdf);
        }
    }

    class imagesToPdf
    {
        public void Convert()
        {
            string path = "../../doc/";
            string filename2 = path + "Rough Draft.png";
            PdfDocument document = new PdfDocument();
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            DrawImage(gfx, filename2, 0, 0, 50, 50);
            document.Save("../../doc/png.pdf");

        }

        private void DrawImage(XGraphics gfx, string jpegPath, int x, int y, int width, int height)
        {
            // Gets Xobject from image 
            XImage image = XImage.FromFile(jpegPath);
            gfx.DrawImage(image, x, y, width, height);
        }
    }
}
