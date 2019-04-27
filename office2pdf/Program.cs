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
            if (args.Length == 2)
            {
                inputPath = args[0];
                outputPath = args[1];

                int dotIndex = inputPath.LastIndexOf(".");
                if (dotIndex > 0 && dotIndex < inputPath.Length - 1)
                {
                    fileType = "-" + inputPath.Substring(dotIndex + 1);
                }
            }
            else if (args.Length == 3)
            {
                fileType = args[0];
                inputPath = args[1];
                outputPath = args[2];
            }
            else
            {
                help();
            }

            switch (fileType)
            {
                case "-ppt":
                case "-pptx":
                    ppt2pdf(inputPath, outputPath);
                    break;
                case "-word":
                case "-doc":
                case "-docx":
                    word2pdf(inputPath, outputPath);
                    break;
                case "-excel":
                case "-xls":
                case "-xlsx":
                    excel2pdf(inputPath, outputPath);
                    break;
                default:
                    help();
                    return;
            }

        }

        static void help()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error, invalid usage!");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Usage: office2pdf.exe  [-ppt|-word|-excel]  inputPath  outputPath");
            Console.ResetColor();
            Environment.Exit(1);
        }

        static void ppt2pdf(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();

                app.Presentations.Open(
                    inputPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoFalse)
                .SaveAs(
                    outputPath,
                    Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }

        static void word2pdf(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();

                app.Documents.Open(
                    inputPath,
                    // Confirm Conversion
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    // Read Only
                    Microsoft.Office.Core.MsoTriState.msoTrue)
                .SaveAs(
                    outputPath,
                    Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }

        static void excel2pdf(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                app.Workbooks.Open(
                    inputPath,
                    Microsoft.Office.Core.MsoTriState.msoFalse,
                    Microsoft.Office.Core.MsoTriState.msoTrue)
                .ExportAsFixedFormat(
                    Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                    outputPath
                    );
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }

        static void ppt2jpg(string inputPath, string outputPath)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();

                app.Presentations.Open(
                    inputPath,
                    Microsoft.Office.Core.MsoTriState.msoTrue, // ReadOnly
                    Microsoft.Office.Core.MsoTriState.msoFalse, // Untitled
                    Microsoft.Office.Core.MsoTriState.msoFalse) // WithWindow
                .SaveAs(
                    outputPath,
                    Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsJPG);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }
    }
}
