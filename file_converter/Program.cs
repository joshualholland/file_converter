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
            // Convert doc to pdf
            string path = "../../doc/";
            string fileName1 = path + "resume.docx";
            Document doc = new Document(fileName1);
            doc.Save(path + "DocumentToPdf.pdf", SaveFormat.Pdf);
        }

        
    }

    class jpgToPdf
    {
        private void Convert()
        {

        }

        void DrawImage(XGraphics gfx, string jpegPath, int x, int y, int width, int height)
        {
            XImage image = XImage.FromFile(jpegPath);
            gfx.DrawImage(image, x, y, width, height);
        }
    }
}
