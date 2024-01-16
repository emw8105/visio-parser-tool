using Microsoft.VisualBasic.FileIO;
using System;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Diagnostics;
using System.ComponentModel;
using System.Diagnostics.Metrics;
using System.Reflection.Metadata;
using System.IO.Compression;
using System.IO.MemoryMappedFiles;
using ArchitectConvert.ConsoleHost;
using System.Drawing;
using System.Text;
using System.Collections.Generic;
using System.Linq.Expressions;
using static System.Net.Mime.MediaTypeNames;

// Written by Evan Wright

namespace VisioParse.ConsoleHost
{
    class VisioParser
    {
        private static readonly XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main"; // needed for queries to match elements, use for all xml parsing

        static void Main(string[] args)
        {
            // Use callflow handler to manage file based setup and modifications, including inserting the desired file name to parse
            CallflowHandler callflow = new CallflowHandler();

            // Configuration options
            callflow.Config.ConfigurationSetup();

            // extract and read xml file contents
            try
            {
                Console.WriteLine("extracting file to " + callflow.ExtractPath);
                ZipFile.ExtractToDirectory(callflow.ZipPath, callflow.ExtractPath); // convert given visio file to xml components
                Console.WriteLine("finished extraction, parsing components...");

                // first parse "pages.xml" to find page count and page names
                XDocument pagesXml;
                using (XmlTextReader documentReader = new XmlTextReader(callflow.ExtractPath + @"\visio\pages\pages.xml"))
                {
                    pagesXml = XDocument.Load(documentReader);
                }

                var xPages = pagesXml.Descendants(ns + "Page");
                var pageCount = xPages.Count();
                var pathCountTotal = 0;
                var minPathCountTotal = 0;

                Console.WriteLine($"Total number of pages: {pageCount}");
                callflow.PageInfoFile.WriteLine($"Total number of pages: {pageCount}");

                XDocument xmlDoc; // used for current xml being parsed

                Page[] pageList = new Page[pageCount];
                var multiPageGraph = new DirectedMultiGraph<VertexShape, EdgeShape>(); // used for multi-flow parsing (off-page references)

                // pages start at page1 and go up to and including the page count
                int i = 1;
                foreach (var page in xPages)
                {
                    var pageName = page?.Attribute("Name")?.Value;
                    Console.WriteLine($"Parsing page {i}: {pageName}");
                    callflow.PageInfoFile.WriteLine($"\n----Page {i}: {pageName}----");

                    using (XmlTextReader reader = new XmlTextReader(callflow.ExtractPath + @"\visio\pages\page" + i + ".xml"))
                    {
                        xmlDoc = XDocument.Load(reader);
                    }

                    var graph = GraphBuilder.BuildGraph(xmlDoc, callflow, i, pageName);
                    MergePageGraph(graph, multiPageGraph); // combine all pages into one graph to parse between off-page references

                    pageList[i - 1] = BuildPage(graph, page, i);

                    // print out the graph data parsed from the page
                    callflow.PageInfoFile.WriteLine($"Page {i} has {graph.Vertices.Count} vertices and {graph.NumberOfEdges} edges");
                    callflow.PageInfoFile.WriteLine($"\nGraph notation of page {i}:");
                    PrintPageInformation(graph, callflow.PageInfoFile);

                    i++;
                }

                Console.WriteLine("Generating permutations...");
                var pathSet = GetAllPermutations(multiPageGraph, callflow.PathOutputFile, callflow.Config, pageList, i);
                var pathCount = pathSet.Count;

                // find the minimum paths for test cases
                if (pathSet.Count > 0)
                {
                    var minPathSet = GetMinimumPaths(multiPageGraph, pathSet, callflow.MinPathOutputFile);
                    var minPathCount = minPathSet.Count;

                    callflow.MinPathOutputFile.WriteLine($"Minimum number of paths on Page {i} to cover all cases: {minPathCount}");
                    callflow.PageInfoFile.WriteLine($"Number of paths to test: {pathCount}\n");
                    Console.WriteLine($"Number of paths to test: {pathCount}\n");
                    pathCountTotal += pathCount;
                    minPathCountTotal += minPathCount;
                }


                Console.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
                Console.WriteLine($"Total minimum number of paths to test: {minPathCountTotal}");
                callflow.PathOutputFile.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
                callflow.PathOutputFile.WriteLine($"Total minimum number of paths to test: {minPathCountTotal}");

                callflow.PageInfoFile.Flush();
                callflow.PageInfoFile.Close();
                callflow.PathOutputFile.Flush();
                callflow.PathOutputFile.Close();

                callflow.ExecutionCleanup();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");

                callflow.PageInfoFile.Flush();
                callflow.PageInfoFile.Close();
                callflow.PathOutputFile.Flush();
                callflow.PathOutputFile.Close();
            }
        }

        static void MergePageGraph(DirectedMultiGraph<VertexShape, EdgeShape> graph, DirectedMultiGraph<VertexShape, EdgeShape> multiPageGraph)
        {
            foreach (var vertex in graph.Vertices)
            {
                multiPageGraph.AddVertex(vertex);
            }
            foreach (var edge in graph.Edges)
            {
                multiPageGraph.Add(graph.Vertices.First(v => v.Id == edge.FromShape), graph.Vertices.First(v => v.Id == edge.ToShape), edge);
            }
        }

        static Page BuildPage(DirectedMultiGraph<VertexShape, EdgeShape> graph, XElement page, int pageNum) =>
            new Page()
            {
                Graph = graph,
                Id = page.Attribute("ID").Value,
                Number = pageNum,
                Name = page.Attribute("Name").Value,
            };

        static void GetSpecifiedNodes(DirectedMultiGraph<VertexShape, EdgeShape> graph, string nodeOption, string? startNodeContent, string? endNodeContent, out IEnumerable<VertexShape> startNodes, out IEnumerable<VertexShape> endNodes)
        {
            switch (nodeOption.Substring(0, 1))
            {
                case "1":
                    Console.WriteLine($"Parsing generated graph to find start nodes with MasterID: {startNodeContent}, and end nodes with MasterID: {endNodeContent}");
                    startNodes = graph.Vertices.Where(vertex => vertex.MasterId == startNodeContent && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    endNodes = graph.Vertices.Where(vertex => vertex.MasterId == endNodeContent && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);
                    break;
                case "2":
                    Console.WriteLine($"Parsing generated graph to find start nodes with text: {startNodeContent}, and end nodes with text: {endNodeContent}");
                    startNodes = graph.Vertices.Where(vertex => vertex.Text == startNodeContent && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    endNodes = graph.Vertices.Where(vertex => vertex.Text == endNodeContent && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);

                    break;
                default: // default case is to parse indiscriminately by checking which nodes have no incoming edges but still have outgoing edges
                    startNodes = graph.Vertices.Where(vertex => graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    endNodes = graph.Vertices.Where(vertex => graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);
                    break;
            }
        }

        static List<List<VertexShape>> GetAllPermutations(DirectedMultiGraph<VertexShape, EdgeShape> graph, StreamWriter file, Configuration config, Page[] pageList, int pageNum)
        {
            GraphBuilder.ConnectReferenceShapes(graph, config.NodeOption, config.StartOffPageContent, config.EndOffPageContent, config.CheckpointContent, pageList);

            // first get the start and end nodes based on user specification
            IEnumerable<VertexShape>? startNodes;
            IEnumerable<VertexShape>? endNodes;
            GetSpecifiedNodes(graph, config.NodeOption, config.StartNodeContent, config.EndNodeContent, out startNodes, out endNodes);

            Console.Write("Start nodes: ");
            startNodes.ToList().ForEach(node => Console.Write($"{node.Id}, "));
            Console.Write("\nEnd nodes: ");
            endNodes.ToList().ForEach(node => Console.Write($"{node.Id}, "));
            Console.WriteLine();

            file.Write("Start nodes: ");
            startNodes.ToList().ForEach(node => file.Write($"{node.Id}, "));
            file.Write("\nEnd nodes: ");
            endNodes.ToList().ForEach(node => file.Write($"{node.Id}, "));
            file.WriteLine();

            
            // list containing the paths containing each node
            var allPaths = new List<List<VertexShape>>();

            int numPaths = 0;

            // find every path from every start node to every end node
            foreach (var startNode in startNodes)
            {
                foreach (var endNode in endNodes)
                {
                    var permutations = FindPermutations(graph, startNode, endNode);
                    allPaths.AddRange(permutations);
                    numPaths += permutations.Count;
                }
            }

            if (allPaths.Count > 0)
            {
                

                file.WriteLine("Paths:");
                PrintPathInformation(allPaths, file, "Id");

                //var minimalPathSet = GetMinimumPaths(graph, allPaths, file);

                file.WriteLine($"\nTotal number of paths on Page {pageNum}: {numPaths}");
                //file.WriteLine($"Minimum number of paths on Page {pageNum} to cover all cases: {minimalPathSet.Count}");

            }
            else
            {
                file.WriteLine($"No test cases to be generated for this page");
            }

            return allPaths;
        }

        // method for finding all of the paths between every starting vertex to every ending vertex
        static List<List<VertexShape>> FindPermutations(DirectedMultiGraph<VertexShape, EdgeShape> graph, VertexShape startNode, VertexShape endNode)
        {
            List<List<VertexShape>> permutationPath = new List<List<VertexShape>>();
            HashSet<VertexShape> visited = new HashSet<VertexShape>();
            List<VertexShape> currentPath = new List<VertexShape>();

            DFS(startNode, endNode);
            return permutationPath;

            void DFS(VertexShape currentNode, VertexShape destinationNode)
            {
                visited.Add(currentNode);
                currentPath.Add(currentNode);

                if (currentNode == destinationNode)
                {
                    // reached the destination, record the current path
                    permutationPath.Add(new List<VertexShape>(currentPath));
                }
                else
                {
                    foreach (var neighbor in graph.GetChildren(currentNode))
                    {
                        if (!visited.Contains(neighbor))
                        {
                            // recursively explore the neighbor nodes
                            DFS(neighbor, destinationNode);
                        }
                    }
                }

                // backtrack
                visited.Remove(currentNode);
                currentPath.RemoveAt(currentPath.Count - 1);
            }
        }

        // this algorithm is much faster but is lazy, will result in approximately 2x as many minimum tests which defeats the purpose
        //static List<List<VertexShape>> GetMinimumPaths(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<List<VertexShape>> allPaths, StreamWriter file)
        //{
        //    // get unique edges from the graph and add them to the covered edges as the minimal paths become concrete
        //    var uniqueEdges = graph.Edges.Distinct().ToList();
        //    HashSet<EdgeShape> coveredEdges = new HashSet<EdgeShape>();

        //    // list of minimal paths covering all unique edges
        //    var minimalPathSet = new List<List<VertexShape>>();

        //    // loop through paths and add to minimalPathSet if it covers a unique edge
        //    foreach (var path in allPaths)
        //    {
        //        var pathCoversUniqueEdge = false;

        //        for (int i = 0; i < path.Count - 1; i++)
        //        {
        //            VertexShape source = path[i];
        //            VertexShape target = path[i + 1];

        //            // find the edge from source to target
        //            if (graph.GetEdges(source, target).Any(edge => uniqueEdges.Contains(edge) && !coveredEdges.Contains(edge)))
        //            {
        //                pathCoversUniqueEdge = true;
        //                coveredEdges.UnionWith(graph.GetEdges(source, target));
        //            }
        //        }

        //        if (pathCoversUniqueEdge)
        //        {
        //            minimalPathSet.Add(path);
        //        }
        //    }

        //    // print the minimal paths
        //    file.WriteLine("\nMinimal paths covering all unique edges:");
        //    PrintPathInformation(minimalPathSet, file, "Id");
        //    file.WriteLine("\nText for minimum paths:");
        //    PrintPathInformation(minimalPathSet, file, "Text");

        //    return minimalPathSet;
        //}

        // this algorithm takes up 95% of runtime and was possibly 50% of the total brainpower
        static List<List<VertexShape>> GetMinimumPaths(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<List<VertexShape>> allPaths, StreamWriter file)
        {
            // get unique edges from the graph and add them to the covered edges as the minimal paths become concrete
            var uniqueEdges = graph.Edges.Distinct().ToList();
            // set to keep track of covered edges
            HashSet<EdgeShape> coveredEdges = new HashSet<EdgeShape>();

            // list of minimal paths covering all unique edges
            var minimalPathSet = new List<List<VertexShape>>();

            // while there are still uncovered edges
            while (coveredEdges.Count < uniqueEdges.Count)
            {
                // sort paths by the number of new uncovered edges
                allPaths.Sort((path1, path2) =>
                {
                    int uncoveredEdgesCount1 = CountUncoveredEdges(graph, path1, uniqueEdges, coveredEdges);
                    int uncoveredEdgesCount2 = CountUncoveredEdges(graph, path2, uniqueEdges, coveredEdges);
                    return uncoveredEdgesCount2.CompareTo(uncoveredEdgesCount1);
                });

                // select the path with the maximum number of new uncovered edges
                var path = allPaths.FirstOrDefault(p => CountUncoveredEdges(graph, p, uniqueEdges, coveredEdges) > 0);

                if (path != null)
                {
                    // add the selected path to the set of selected paths
                    minimalPathSet.Add(path);
                    // update the set of covered edges
                    coveredEdges.UnionWith(GetEdgesFromPath(graph, path));
                }
                else
                {
                    // if there is no path that covers new edges, break the loop
                    break;
                }
            }

            // print the minimal paths
            file.WriteLine("\nMinimal paths covering all unique edges:");
            PrintPathInformation(minimalPathSet, file, "Id");
            file.WriteLine("\nText for minimum paths:");
            PrintPathInformation(minimalPathSet, file, "Text");

            return minimalPathSet;
        }

        static int CountUncoveredEdges(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<VertexShape> path, List<EdgeShape> uniqueEdges, HashSet<EdgeShape> coveredEdges)
            => GetEdgesFromPath(graph, path).Count(edge => uniqueEdges.Contains(edge) && !coveredEdges.Contains(edge));

        static List<EdgeShape> GetEdgesFromPath(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<VertexShape> path)
        {
            var edges = new List<EdgeShape>();

            for (int i = 0; i < path.Count - 1; i++)
            {
                VertexShape source = path[i];
                VertexShape target = path[i + 1];

                // find the edge from source to target
                edges.AddRange(graph.GetEdges(source, target).Where(edge => !edges.Contains(edge)));
            }

            return edges;
        }

        

        static void PrintPageInformation(DirectedMultiGraph<VertexShape, EdgeShape> graph, StreamWriter file)
        {
            file.WriteLine("Vertices:");
            foreach (var vertex in graph.Vertices)
            {
                file.WriteLine($"Vertex: {vertex.Id}, InDegree: {graph.GetInDegree(vertex)}, OutDegree: {graph.GetOutDegree(vertex)}, Text: {vertex.Text}");
            }
            file.WriteLine("\nEdges:");
            foreach (var edge in graph.Edges)
            {
                file.WriteLine($"Edge: {edge.Id} connects vertex {edge.FromShape} to vertex {edge.ToShape}");
            }
            file.WriteLine();
        }

        static void PrintPathInformation(List<List<VertexShape>> pathSet, StreamWriter file, string property)
        {
            int count = 1;
            switch (property)
            {
                case "Id":
                    foreach (var path in pathSet)
                    {
                        file.Write($"{count}. ");
                        foreach (var node in path)
                        {
                            file.Write($"{node.Id} -> ");
                        }
                        file.WriteLine();
                        count++;
                    }
                    break;
                case "Text":
                    foreach (var path in pathSet)
                    {
                        file.Write($"{count}. ");
                        foreach (var node in path)
                        {
                            file.Write($"{node.Text} -> ");
                        }
                        file.WriteLine();
                        count++;
                    }
                    break;
                default:
                    Console.WriteLine($"Unrecognized property to print: {property}");
                    break;
            }
        }
    }
}