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


        // example files from solution:
        // note that USPS flows don't have all references attached and use extra connections for start/end nodes, dont always use the same end node
        // DCW and CE use return statements

        // for basic impleplementation and testing: Basic.vsdx,
        // for a simple test: Simple Test.vsdx

        // best testcase if time permits: Parsable DCWater_IVR_Callflow v5.0 (Post Go -Live Kubra Replacement).vsdx
        // best testcase on short notice: Parsable Inbound Routing Design v1.18.vsdx
        public CallflowHandler()
        {
            // set this value prior to running program based on the desired visio
            //FileName = "Simple Test.vsdx";

            // get all .vsdx files in the Documents directory and print them out for the user to choose from
            var files = Directory.GetFiles("Documents", "*.vsdx");

            // if there's only one file then automatically select it
            if (files.Length == 1)
            {
                FileName = System.IO.Path.GetFileName(files[0]);
                Console.WriteLine($"Only one Visio file found: {FileName}. Automatically selected.");
            }
            else // else ask the user to select a file
            {
                // print the files for the user
                for (int i = 0; i < files.Length; i++)
                {
                    Console.WriteLine($"{i + 1}. {System.IO.Path.GetFileName(files[i])}");
                }

                // ask the user to select a file
                int index;
                do
                {
                    Console.Write("Enter the number of the Visio file you want to parse: ");
                    var input = Console.ReadLine();
                    if (int.TryParse(input, out index) && index > 0 && index <= files.Length)
                    {
                        FileName = System.IO.Path.GetFileName(files[index - 1]);
                    }
                    else
                    {
                        Console.WriteLine("Invalid selection. Please try again.");
                        index = 0; // reset index to an invalid value
                    }
                } while (index <= 0 || index > files.Length);
            }

            Console.WriteLine($"Currently parsing: {FileName}");
            var fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(FileName);
            Console.WriteLine("Current directory: " + System.IO.Directory.GetCurrentDirectory());
            Path = "";
            Console.WriteLine(Path);
            ZipPath = Path + @"Documents\" + FileName; // path to get the zipped file i.e. the visio
            ExtractPath = Path + @"extracted"; // path to extract to within console host file

            // create output files
            string pageInfoFilePath = System.IO.Path.Combine(Path, "pageInfo_" + fileNameWithoutExtension + ".txt");
            File.WriteAllText(pageInfoFilePath, "Beginning output\n");
            PageInfoFile = File.AppendText(pageInfoFilePath);

            string pathsFilePath = System.IO.Path.Combine(Path, "paths_" + fileNameWithoutExtension + ".txt");
            File.WriteAllText(pathsFilePath, "Beginning output\n");
            PathOutputFile = File.AppendText(pathsFilePath);

            pathsFilePath = System.IO.Path.Combine(Path, "minPaths_" + fileNameWithoutExtension + ".txt");
            File.WriteAllText(pathsFilePath, "Beginning output\n");
            MinPathOutputFile = File.AppendText(pathsFilePath);

            if (!File.Exists(ZipPath))
            {
                Console.WriteLine($"The Visio file was not found in the Documents folder: {FileName}");
                Console.WriteLine("First, enter a Visio file name into the FileName property in the CallflowHandler constructor within CallflowHandler.cs");
                Console.WriteLine("The Visio must be contained within the Documents folder");
                Console.WriteLine("Check if the entered Visio file name is correct and ensure that it is set to 'Copy if newer' in the Visual Studio properties");
                Console.WriteLine("Press any key to exit, please try again");
                Console.ReadLine();
            }
            else
            {
                // generate a configuration for this parsing with options selected by the user
                Config = new Configuration();
                Config.ConfigurationSetup();

                // extract the XML contents
                Console.WriteLine("extracting file to " + ExtractPath);
                try
                {
                    // first delete any existing extract path if execution was paused halfway from a previous run
                    if (Directory.Exists(ExtractPath))
                    {
                        Directory.Delete(ExtractPath, true);
                    }
                    ZipFile.ExtractToDirectory(ZipPath, ExtractPath); // convert given visio file to xml components
                    Console.WriteLine("finished extraction, parsing components...");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Please make sure to close the Visio before parsing and ensure that it has been set to 'Copy if newer' in Visual Studio: ", ex);
                    Console.ReadLine(); // put here for inspection
                }
            }
        }

        /// <summary>
        /// Gets the "pages.XML" file which contains various information about each page such as the page count
        /// </summary>
        /// <returns>pages.xml as an XDocument</returns>
        public XDocument GetPagesXML()
        {
            using (XmlTextReader documentReader = new XmlTextReader(ExtractPath + @"\visio\pages\pages.xml"))
            {
                return XDocument.Load(documentReader);
            }
        }

        /// <summary>
        /// Gets a single page of XML
        /// </summary>
        /// <param name="pageNum">The page number to get</param>
        /// <returns>The desired page as an XDocument</returns>
        public XDocument GetPageXML(int pageNum)
        {
            using (XmlTextReader reader = new XmlTextReader(ExtractPath + @"\visio\pages\page" + pageNum + ".xml"))
            {
                return XDocument.Load(reader);
            }
        }

        /// <summary>
        /// Provides the user with choices to deal with leftover files for quick deletion
        /// </summary>
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

                        string newZipPath = Path + @"Modified" + FileName + ".zip";
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
                            }
                            else
                            {
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

        /// <summary>
        /// Handles file output flushing and closing
        /// </summary>
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
