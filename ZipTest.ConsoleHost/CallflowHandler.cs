using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;

namespace VisioParse.ConsoleHost
{
    public class CallflowHandler
    {
        // example files from solution:
        // ECC IVR Call Flow V104.1_updated.vsdx, USPS ITHD IVR LiteBlue MFA Ticket 8_30_2023.vsdx,
        // USPS_GCX_NMCSC_IVRCallFlow_006 (1).vsdx, Basic.vsdx, Inbound Routing Design v1.18.vsdx

        // set these values prior to running program depending on user's directory
        public string Path = @"C:\Users\ewright\source\repos\ZipTest.ConsoleHost\";
        public string FileName = "Improved Inbound Routing Design v1.18.vsdx";
        public string YamlFileName = "Select Health Routing Research_v11-0.yaml";

        // generated at runtime using constructor
        public string ZipPath = string.Empty;
        public string ExtractPath = string.Empty;
        public string YamlFilePath = string.Empty;
        public StreamWriter PageInfoFile;
        public StreamWriter PathOutputFile;

        public CallflowHandler()
        {
            YamlFilePath = Path + @"Documents\" + YamlFileName;
            ZipPath = Path + @"Documents\" + FileName; // path to get the zipped file i.e. the visio
            ExtractPath = Path + "extracted"; // path to extract to within console host file

            // create output files
            string pageInfoFilePath = System.IO.Path.Combine(Path, "pageInfo_" + FileName + ".txt");
            File.WriteAllText(pageInfoFilePath, "Beginning output\n");
            PageInfoFile = File.AppendText(pageInfoFilePath);

            string pathsFilePath = System.IO.Path.Combine(Path, "paths_" + FileName + ".txt");
            File.WriteAllText(pathsFilePath, "Beginning output\n");
            PathOutputFile = File.AppendText(pathsFilePath);
        }
        public void ExecutionCleanup()
        {
            // rezip the file, change the extension to .vsdx
            Console.WriteLine("\nFinished parsing, check the output files for more info:\n1. Rezip the extracted files to inspect Shape ID's for permutation comparison" +
                "\n2. Exit without deleting extracted files" +
                "\n3. Delete the extracted/copied files");
            string? input = Console.ReadLine(); // readline to allow for inspecting the file and delete when done

            switch (input)
            {
                case "1":
                    string newZipPath = Path + @"\Modified" + FileName + ".zip";
                    ZipFile.CreateFromDirectory(ExtractPath, newZipPath);

                    // change the file extension and move it to the zip file to create a visio file
                    string newVisioPath = System.IO.Path.ChangeExtension(newZipPath, ".vsdx");
                    File.Move(newZipPath, newVisioPath);

                    Console.WriteLine("\nfinished zipping, type 'd' to delete the visio reconstructed from the original data and the extracted files");
                    input = Console.ReadLine(); // readline to allow for inspecting the file and delete when done
                    if (input == "d")
                    {
                        File.Delete(newVisioPath);
                        Directory.Delete(ExtractPath, true);
                        Console.WriteLine("\nExtracted visio and files have been deleted");
                    }
                    break;
                case "3": // note that the "extracted" folder will need to be moved or deleted by hand so that the program can run again
                    Directory.Delete(ExtractPath, true);
                    Console.WriteLine("\nDirectory has been deleted");
                    break;
                default: // case 2 is the default case, delete the extra files to allow for running the program again with no issue
                    Directory.Delete(ExtractPath, true);
                    Console.WriteLine("\nDirectory has been deleted");
                    break;
            }
        }
    }
}
