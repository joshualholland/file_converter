using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using Aspose.Words;
using PdfSharp.Pdf;
using PdfSharp.Drawing;

namespace file_converter
{
    class Program
    {
        static void Main(string[] args)
        {
            docToPdf word = new docToPdf();
            word.Convert();
            jpgToPdf image = new jpgToPdf();
            image.Convert();
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

    class jpgToPdf
    {
        public void Convert()
        {
            string path = "../../doc/";
            string filename2 = path + "me.jpg";
            PdfDocument document = new PdfDocument();
            PdfPage page = document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(page);
            DrawImage(gfx, filename2, 0, 0, 50, 50);
            document.Save("../../doc/image.pdf");

        }

        private void DrawImage(XGraphics gfx, string jpegPath, int x, int y, int width, int height)
        {
            XImage image = XImage.FromFile(jpegPath);
            gfx.DrawImage(image, x, y, width, height);
        }
    }
}
