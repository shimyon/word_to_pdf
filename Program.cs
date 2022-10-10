//using Microsoft.Office.Interop.Word;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.DocToPDFConverter;
using Syncfusion.Pdf;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Words;
using WatermarkType = Aspose.Words.WatermarkType;
using GrapeCity.Documents.Word;
using Microsoft.Office.Interop.Word;
using System.IO;

namespace ReadWordFile
{
    class Program
    {

        static void Main(string[] args)
        {

    

            genFile();
          
        }

        static void genFile()
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();

            // C# doesn't have optional arguments so we'll need a dummy value
            object oMissing = System.Reflection.Missing.Value;

            // Get list of Word files in specified directory
            //DirectoryInfo dirInfo = new DirectoryInfo(@"D:\docfiles");
            //FileInfo[] wordFiles = dirInfo.GetFiles("*.doc");
            Console.Write("Enter File Path: ");
            var val = Console.ReadLine();
            FileInfo singleFile = new FileInfo(val);
            Console.WriteLine("Converting...");

            word.Visible = false;
            word.ScreenUpdating = false;

           
                // Cast as Object for word Open method
                Object filename = (Object)singleFile.FullName;

                // Use the dummy value as a placeholder for optional arguments
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);
                doc.Activate();

                object outputFileName = singleFile.FullName.Replace(".doc", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                    ref fileFormat, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.                
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref oMissing, ref oMissing);
                doc = null;
            //}

        ((_Application)word).Quit(ref oMissing, ref oMissing, ref oMissing);
            word = null;
        }
    }
}
