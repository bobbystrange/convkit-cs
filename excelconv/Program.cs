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

            Microsoft.Office.Interop.Excel.XlFixedFormatType? type = null;
            Microsoft.Office.Interop.Excel.XlFileFormat? format = null;
            switch (fileType)
            {
                case "-pdf":
                    type = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF;
                    break;
                case "-xps":
                    type = Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypeXPS;
                    break;
                case "-csv":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSV;
                    break;
                case "-csv-mac":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVMac;
                    break;
                case "-csv-dos":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVMSDOS;
                    break;
                case "-csv-windows":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows;
                    break;
                case "-xml":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlXMLSpreadsheet;
                    break;
                case "-xls":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook;
                    break;
                case "-xlsx":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLStrictWorkbook;
                    break;
                case "-html":
                    format = Microsoft.Office.Interop.Excel.XlFileFormat.xlHtml;
                    break;
                default:
                    help();
                    return;
            }

            if (type != null)
                exportAsFixedFormat(inputPath, outputPath, (Microsoft.Office.Interop.Excel.XlFixedFormatType)type);
            else
                saveAs(inputPath, outputPath, (Microsoft.Office.Interop.Excel.XlFileFormat)format);

        }

        static void help()
        {
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Error, invalid usage!");
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Usage: excelconv.exe  -pdf|-xps|-csv|-xml|-html|-xls|-xlsx  inputPath  outputPath");
            Console.ResetColor();
            Environment.Exit(1);
        }

        static void exportAsFixedFormat(string inputPath, string outputPath, Microsoft.Office.Interop.Excel.XlFixedFormatType type)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                app.Workbooks.Open(inputPath)
                    .ExportAsFixedFormat(type, outputPath);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }

        static void saveAs(string inputPath, string outputPath, Microsoft.Office.Interop.Excel.XlFileFormat format)
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException(string.Format("The specified file {0} does not exist.", inputPath), inputPath);
            }

            try
            {
                Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();

                app.Workbooks.Open(inputPath)
                    .SaveAs(outputPath, format);
                app.Quit();
            }
            catch (Exception e)
            {
                throw new Exception(string.Format("Unable to convert {0} to {1}", inputPath, outputPath), e);
            }
        }



    }
}
