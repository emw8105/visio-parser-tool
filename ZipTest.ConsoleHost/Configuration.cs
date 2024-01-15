using ArchitectConvert.ConsoleHost;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VisioParse.ConsoleHost
{
    public class Configuration
    {
        // generated at runtime from user input to serve as configuration options
        // in the future, most of these can be made into a general template (i.e. use Master ID 4 for start/end nodes, etc.)

        public string? NodeOption = string.Empty; // 2 digits representing config options, first is for start/end nodes, second is for multi-flow parsing
        // basic config for start/end nodes
        public string? StartNodeContent = string.Empty; // either text or Master ID representing the value to use when considering start nodes
        public string? EndNodeContent = string.Empty; // either text or Master ID representing the value to use when considering start nodes
        // advanced config for multi-flow parsing, checkpoints refer to shapes which are not hyperlinked and appear at various points in other flows
        public string? StartOffPageContent = string.Empty; // the Master ID of start node off-page reference shapes (i.e. "FROM: X" nodes)
        public string? EndOffPageContent = string.Empty; // the Master ID of end node off-page reference shapes (i.e. "TO: X" nodes)
        public string? CheckpointContent = string.Empty; // the Master ID of off-page reference shapes, the text has a format

        public void ConfigurationSetup()
        {
            Console.WriteLine("Please select a menu option:" +
                //"\n1. Convert a Genesys Architect flow into Visio" +
                "\n2. Parse Visio information with specified start and end nodes  (more precise)" + // can ask user to use text or a master ID
                "\n3. Parse Visio information indiscriminately  (less precise)");
            string? menuChoice = Console.ReadLine();
            switch (menuChoice)
            {
                case "2":
                    Console.WriteLine("Choose your ideal method of determining start nodes based on your Visio structure:" +
                        "\n1. Parse based on a Master ID (a specific shape used)" +
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
                    NodeOption = "0";
                    break;
            }

            Console.WriteLine("Please select from the additional configuration options:" +
                "\n1. Multi-flow parsing using off-page references" +
                "\n2. Multi-flow parsing using checkpoints" +
                "\n3. Multi-flow parsing using both off-page references and checkpoints" +
                "\n4. No multi-flow parsing");
            Console.WriteLine("Note: The off-page references must all be hyperlinked properly and checkpoints must follow the following format:" +
                "pageName: checkPointIdentifier (ex: Legend: A)");
            menuChoice = Console.ReadLine();
            switch (menuChoice)
            {
                case "1":
                    NodeOption += "1";
                    Console.WriteLine("Please enter the Master ID for the starting off-page reference shapes (i.e. 'FROM: X' nodes");
                    StartOffPageContent = Console.ReadLine();
                    Console.WriteLine("Please enter the Master ID for the ending off-page reference shapes (i.e. 'TO: X' nodes");
                    EndOffPageContent = Console.ReadLine();
                    break;

                case "2":
                    NodeOption += "2";
                    Console.WriteLine("Please enter the Master ID for checkpoint/on-page reference shapes");
                    CheckpointContent = Console.ReadLine();
                    break;

                case "3":
                    NodeOption += "3";
                    Console.WriteLine("Please enter the Master ID for the starting off-page reference shapes (i.e. 'FROM: X' nodes");
                    StartOffPageContent = Console.ReadLine();
                    Console.WriteLine("Please enter the Master ID for the ending off-page reference shapes (i.e. 'TO: X' nodes");
                    EndOffPageContent = Console.ReadLine();
                    Console.WriteLine("Please enter the Checkpoint ID");
                    CheckpointContent = Console.ReadLine();
                    break;

                default: // Case 4 is the default (No multi-flow parsing)
                    NodeOption += "4";
                    Console.WriteLine("Multi-flow parsing disabled, off-page references and checkpoints will be treated as regular vertices and paths won't continue between pages");
                    break;
            }
        }
    }
}
