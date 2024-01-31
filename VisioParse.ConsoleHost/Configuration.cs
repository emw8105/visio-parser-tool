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
                "\n1. Parse Visio information with specified start and end nodes  (more precise)" + // can ask user to use text or a master ID
                "\n2. Parse Visio information indiscriminately  (less precise)");
            string? menuChoice = Console.ReadLine();
            switch (menuChoice)
            {
                case "1":
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
                "\n2. Multi-flow parsing using on-page references" +
                "\n3. Multi-flow parsing using both off-page references and on-page references" +
                "\n4. No multi-flow parsing");
            Console.WriteLine("Note: The off-page references must all be hyperlinked properly and on-page references must contain only the text of their identifiers (ex: A)");
            menuChoice = Console.ReadLine();
            switch (menuChoice)
            {
                case "1":
                    NodeOption += "1";
                    Console.WriteLine("Currently, the Master ID does not need to be used to determine off-page references as long as the references are attached");
                    break;

                case "2":
                    NodeOption += "2";
                    Console.WriteLine("Please enter the Master ID for on-page reference shapes");
                    CheckpointContent = Console.ReadLine();
                    break;

                case "3":
                    NodeOption += "3";
                    Console.WriteLine("Currently, the Master ID does not need to be used to determine off-page references as long as the references are attached");
                    Console.WriteLine("Please enter the Master ID for the on-page reference shapes");
                    CheckpointContent = Console.ReadLine();
                    break;

                case "4":
                    NodeOption += "4";
                    Console.WriteLine("Multi-flow parsing disabled, off-page references and checkpoints will be treated as regular vertices and paths won't continue between pages");
                    break;

                default: // Case 4 is the default (No multi-flow parsing)
                    NodeOption += "4";
                    Console.WriteLine("Choice not recognized, multi-flow parsing disabled, off-page references and checkpoints will be treated as regular vertices and paths won't continue between pages");
                    break;
            }
        }
    }
}
