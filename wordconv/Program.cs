using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
namespace ConvKit
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileType = null;
            string inputPath = null;
            string outputPath = null;
            if (args.Length == 3)
            {
                fileType = args[0];
                inputPath = args[1];
                outputPath = args[2];
            }
            else if (args.Length == 2)
            {
                inputPath = args[0];
                outputPath = args[1];
                int dotIndex = outputPath.LastIndexOf(".");
                if (dotIndex > 0 && dotIndex < inputPath.Length - 1)
                {
                    fileType = "-" + outputPath.Substring(dotIndex + 1);
                    outputPath = outputPath.Substring(0, dotIndex);
                }
            }
            else
            {
                help();
                return;
            }

            Microsoft.Office.Interop.Word.WdSaveFormat type;
            switch (fileType)
            {
                case "-pdf":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;
                    break;
                case "-xps":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXPS;
                    break;
                case "-xml":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
                    break;
                case "-xml-flat":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatFlatXML;
                    break;
                case "-xml-document":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatXMLDocument;
                    break;
                case "-html":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatHTML;
                    break;
                case "-doc":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatOpenDocumentText;
                    break;
                case "-docx":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatStrictOpenXMLDocument;
                    break;
                case "-txt":
                case "-text":
                    type = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatUnicodeText;
                    break;
                default:
                    help();
                    return;
            }

            saveAs(inputPath, outputPath, type);

        }

        static void help()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error, invalid usage!");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Usage: wordconv.exe  -pdf|-xps|-xml|-xml-flat|-xml-document|-html|-text|-doc|-docs  inputPath  outputPath");
            Console.ResetColor();
            Environment.Exit(1);
        }

        static void saveAs(string inputPath, string outputPath, Microsoft.Office.Interop.Word.WdSaveFormat type)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                var app = new Microsoft.Office.Interop.Word.Application();

                app.Documents.Open(
                    inputPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse, // Confirm Conversion
                    Microsoft.Office.Core.MsoTriState.msoTrue) // Read Only
                .SaveAs(outputPath, type);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }

    }
}
