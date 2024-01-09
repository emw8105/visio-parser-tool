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

        // generated at runtime from user input
        public string? NodeOption = string.Empty;
        public string? StartNodeContent = string.Empty;
        public string? EndNodeContent = string.Empty;


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
        public void ConfigurationSetup()
        {
            Console.WriteLine("Please select a menu option:" +
                "\n1. Convert a Genesys Architect flow into Visio" +
                "\n2. Parse Visio information with specified start and end nodes  (more precise)" + // can ask user to use text or a master ID
                "\n3. Parse Visio information indiscriminately  (less precise)");
            string? menuChoice = Console.ReadLine();
            switch (menuChoice)
            {
                case "1":
                    // call the class, pass the converter the input .yaml file and the output string
                    // intention is to convert yaml files into visio files, then can use subsequent logic in this project to parse the generated visio
                    FileConverter converter = new FileConverter();
                    string outputPath = Path + "converted";
                    converter.ConvertToVisio(YamlFileName, outputPath);
                    break;
                case "2":
                    Console.WriteLine("Choose your ideal method of determining start nodes based on your Visio structure:" +
                        "\n1. Parse based on a master shape ID (a specific shape used)" +
                        "\n2. Parse based on text (ex: Start)");
                    string? startNodeChoice = Console.ReadLine();
                    switch (startNodeChoice)
                    {
                        case "1":
                            NodeOption = "1"; // represents parsing via master shape ID
                            Console.WriteLine("Please enter the Master ID of the shape to use as starting nodes");
                            StartNodeContent = Console.ReadLine();
                            Console.WriteLine("Please enter the Master ID of the shape to use as ending nodes");
                            EndNodeContent = Console.ReadLine();
                            break;
                        case "2":
                            NodeOption = "2"; // represents parsing via text
                            Console.WriteLine("Please enter the text of the shape to use as starting nodes");
                            StartNodeContent = Console.ReadLine();
                            Console.WriteLine("Please enter the text of the shape to use as starting nodes");
                            EndNodeContent = Console.ReadLine();
                            break;
                        default:
                            NodeOption = "0"; // represents indiscriminate parsing, although it's the default case rather than explicitely evaluated
                            Console.WriteLine("Choice not recognized, parsing indiscriminately instead");
                            break;
                    }
                    break;
                default: // indiscriminate parsing is the default
                    break;
            }
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
