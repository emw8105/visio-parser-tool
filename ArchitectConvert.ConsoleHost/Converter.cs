using System;
using System.IO;
using YamlDotNet.Serialization;

namespace ArchitectConvert.ConsoleHost
{
    public class FileConverter
    {
        public void ConvertToVisio(string inputFile, string outputFile)
        {
            // parse thru architect yaml to generate graph notation
            // then make an excel sheet and put the data in tabular form
            // can open the excel and generate the visio from that?
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 2)
            {
                var inputFile = args[0];
                var outputFile = args[1];

                FileConverter converter = new FileConverter();
                converter.ConvertToVisio(inputFile, outputFile);
            }
            else
            {
                Console.WriteLine("Usage: ArchitectConvert.ConsoleHost <inputFile> <outputFile>");
            }
        }
    }
}