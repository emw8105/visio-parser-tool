using ArchitectConvert.ConsoleHost;
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
        // for basic impleplementation and testing: Basic.vsdx,
        // for a challenge: ECC IVR Call Flow V104.1_updated.vsdx, USPS ITHD IVR LiteBlue MFA Ticket 8_30_2023.vsdx, USPS_GCX_NMCSC_IVRCallFlow_006 (1).vsdx
        // for design practice differences: Inbound Routing Design v1.18.vsdx, Improved Inbound Routing Design v1.18.vsdx
        // extra: DCWater_IVR_Callflow v5.0 (Post Go -Live Kubra Replacement).vsdx
        //      (doesn't work because under the hood, the master id of the starting/ending shapes are different and indiscriminate has hundreds of thousands of paths

        // set these values prior to running program depending on user's directory
        public string Path = @"C:\Users\ewright\source\repos\ZipTest.ConsoleHost\";
        public string FileName = "USPS_GCX_NMCSC_IVRCallFlow_006 (1) - Copy.vsdx";
        public string YamlFileName = "Select Health Routing Research_v11-0.yaml";

        // generated at runtime using constructor
        public string ZipPath = string.Empty;
        public string ExtractPath = string.Empty;
        public string YamlFilePath = string.Empty;
        public StreamWriter PageInfoFile;
        public StreamWriter PathOutputFile;
        public StreamWriter MinPathOutputFile;

        public Configuration Config {get; set;}

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

            pathsFilePath = System.IO.Path.Combine(Path, "minPaths_" + FileName + ".txt");
            File.WriteAllText(pathsFilePath, "Beginning output\n");
            MinPathOutputFile = File.AppendText(pathsFilePath);

            Config = new Configuration();
        }

        public void ExecutionCleanup()
        {
            Console.WriteLine("\nFinished parsing, check the output files for more info:\n" +
                "1. Rezip the extracted files to inspect Shape ID's for permutation comparison\n" +
                "2. Exit without deleting extracted files\n" +
                "3. Delete the extracted/copied files");

            bool valid = false;
            string? input = Console.ReadLine();
            do
            {
                switch (input)
                {
                    // rezip the extracted files into a Visio, allow the user to delete again
                    case "1":

                        string newZipPath = Path + @"\Modified" + FileName + ".zip";
                        ZipFile.CreateFromDirectory(ExtractPath, newZipPath);

                        string newVisioPath = System.IO.Path.ChangeExtension(newZipPath, ".vsdx");
                        File.Move(newZipPath, newVisioPath);

                        // reloops if the deletion is unsuccessful, otherwise any input is valid
                        do
                        {
                            Console.WriteLine("\nFinished zipping, type 'd' to delete the visio reconstructed from the original data and the extracted files");
                            string? deleteInput = Console.ReadLine();
                            if (deleteInput == "d")
                            {
                                try
                                {
                                    File.Delete(newVisioPath);
                                    Directory.Delete(ExtractPath, true);
                                    Console.WriteLine("\nExtracted visio and files have been deleted");
                                    valid = true;
                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Modified Visio file must be closed before deleting. Please close the file and try again.");
                                    continue;
                                };
                                valid = true;
                            }
                        } while (!valid);
                        break;

                    case "2":
                        Console.WriteLine("\nExiting without deleting extracted files.");
                        valid = true;
                        break;

                    case "3":
                        Directory.Delete(ExtractPath, true);
                        Console.WriteLine("\nDirectory has been deleted.");
                        valid = true;
                        break;

                    default:
                        Console.WriteLine("\nInvalid input. Please enter a valid option.");
                        break;
                }
            } while (!valid);
        }
    }
}
