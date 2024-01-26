using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml;
using System.Reflection;

namespace VisioParse.ConsoleHost
{
    /// <summary>
    /// A class that abstracts direct handling of files and directories as well as configuration options
    /// </summary>
    public class CallflowHandler
    {
        // example files from solution:
        // note that USPS flows don't have all references attached and use extra connections for start/end nodes, dont always use the same end node
        // DCW and CE use return statements

        // for basic impleplementation and testing: Basic.vsdx,
        // for a challenge: ECC IVR Call Flow V104.1_updated.vsdx, USPS ITHD IVR LiteBlue MFA Ticket 8_30_2023.vsdx, USPS_GCX_NMCSC_IVRCallFlow_006 (1).vsdx
        // for design practice differences: Inbound Routing Design v1.18.vsdx, Parsable Inbound Routing Design v1.18.vsdx
        // for directory testing: CE_VCC_IVR_Callflow_V5.4.1117.vsdx
        // for a comprehensive test: Comprehensive test.vsdx
        // extra: DCWater_IVR_Callflow v5.0 (Post Go -Live Kubra Replacement).vsdx
        //      (doesn't work because under the hood, the master id of the starting/ending shapes are different and indiscriminate has hundreds of thousands of paths
        // best testcase if time permits: new DCWater_IVR_Callflow v5.0 (Post Go -Live Kubra Replacement).vsdx

        /// <summary>
        /// The name of the visio file to parse which should be set manually
        /// </summary>
        public string FileName { get; set; }

        /// <summary>
        /// The base path to the solution repository
        /// </summary>
        public string Path { get; set; }

        /// <summary>
        /// The path to the documents folder in the repository where the desired Visio can be found and unzipped
        /// </summary>
        public string ZipPath { get; set; }

        /// <summary>
        /// The path to create the extracted folder containing the unzipped XML contents of the Visio
        /// </summary>
        public string ExtractPath { get; set; }

        /// <summary>
        /// An output file storing primarily shape info on pages
        /// </summary>
        public StreamWriter PageInfoFile { get; set; }

        /// <summary>
        /// An output file storing the exhaustive paths
        /// </summary>
        public StreamWriter PathOutputFile { get; set; }

        /// <summary>
        /// An output file storing the minimum paths
        /// </summary>
        public StreamWriter MinPathOutputFile { get; set; }

        /// <summary>
        /// A container for the configuration options desired by the user
        /// </summary>
        public Configuration Config { get; set; }

        public CallflowHandler()
        {
            // set this value prior to running program based on the desired visio
            FileName = "Parsable Inbound Routing Design v1.18.vsdx";

            Path = "";
            Console.WriteLine(Path);
            ZipPath = Path + @"Documents\" + FileName; // path to get the zipped file i.e. the visio
            ExtractPath = Path + @"extracted"; // path to extract to within console host file

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

            // generate a configuration for this parsing with options selected by the user
            Config = new Configuration();
            Config.ConfigurationSetup();

            // extract the XML contents
            Console.WriteLine("extracting file to " + ExtractPath);
            try
            {
                ZipFile.ExtractToDirectory(ZipPath, ExtractPath); // convert given visio file to xml components
                Console.WriteLine("finished extraction, parsing components...");
            }
            catch(Exception ex)
            {
                Console.WriteLine("Please make sure to close the Visio before parsing: ", ex);
            }
        }

        public XDocument GetPagesXML()
        {
            using (XmlTextReader documentReader = new XmlTextReader(ExtractPath + @"\visio\pages\pages.xml"))
            {
                return XDocument.Load(documentReader);
            }
        }

        public XDocument GetPageXML(int pageNum)
        {
            using (XmlTextReader reader = new XmlTextReader(ExtractPath + @"\visio\pages\page" + pageNum + ".xml"))
            {
                return XDocument.Load(reader);
            }
        }

        public void ExecutionCleanup()
        {
            CleanupFiles(); // flush buffer first so that all the text is printed for inspection

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

        public void CleanupFiles()
        {
            PageInfoFile.Flush();
            PageInfoFile.Close();
            PathOutputFile.Flush();
            PathOutputFile.Close();
            MinPathOutputFile.Flush();
            MinPathOutputFile.Close();
        }
    }
}
