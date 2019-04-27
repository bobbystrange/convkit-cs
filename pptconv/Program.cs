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

            Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType type;
            switch (fileType)
            {
                case "-pdf":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF;
                    break;
                case "-xps":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsXPS;
                    break;
                case "-xml":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsXMLPresentation;
                    break;
                case "-jpg":
                case "-jpeg":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsJPG;
                    break;
                case "-png":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPNG;
                    break;
                case "-tif":
                case "-tiff":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsTIF;
                    break;
                case "-bmp":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsBMP;
                    break;
                case "-gif":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsGIF;
                    break;
                case "-mp4":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsMP4;
                    break;
                case "-wmv":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsWMV;
                    break;
                case "-ppt":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation;
                    break;
                case "-pptx":
                    type = Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsStrictOpenXMLPresentation;
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
            Console.WriteLine("Usage: pptconv.exe  -jpeg|-png|-tiff|-bmp|-gif|-mp4|-wmv|-pdf|-xps|-ppt|-pptx inputPath  outputPath");
            Console.ResetColor();
            Environment.Exit(1);
        }

        static void saveAs(string inputPath, string outputPath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType type)
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
