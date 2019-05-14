using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using System.Web;
using PDF_DLL;
using System.Drawing.Text;

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
            PDF_DLL.ConvertToPdf pdf = new ConvertToPdf();
            pdf.Convert(files);

            /* specific file for testing
            string test = "../../doc/Permit APA8429.xls";

            FileStream fs = new FileStream(test, FileMode.Create, FileAccess.ReadWrite);
            PDF_DLL.ConvertFromStream pdf = new ConvertFromStream();
            pdf.ConvertFilestream(fs);*/
        }
    }
}
