﻿using Microsoft.VisualBasic.FileIO;
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
using System.Drawing;
using static System.Runtime.InteropServices.JavaScript.JSType;

// take master id out of off-page references, fix reference algorithm by removing parallel hashmap, update readme, update documentation

namespace VisioParse.ConsoleHost
{
    class Program
    {
        private static readonly XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main"; // needed for queries to match elements, use for all xml parsing

        static void Main(string[] args)
        {
            // use the callflow handler to manage file based setup and modifications, also for configuration management
            // upon construction will set up the output files and configuration settings, then extract the XML files from the desired Visio
            CallflowHandler callflow = new CallflowHandler();

            if (callflow.FileName != string.Empty)
            {
                // use the graph builder to piece together the graph either by using the XML files or information stored in the shapes
                GraphBuilder graphBuilder = new GraphBuilder(callflow);

                // read the extracted XML contents
                try
                {
                    // first parse "pages.xml" to find page count and page names
                    XDocument pagesXml = callflow.GetPagesXML();

                    var xPages = pagesXml.Descendants(ns + "Page");
                    var pageCount = xPages.Count();
                    var pathCountTotal = 0;
                    var minPathCountTotal = 0;
                    Console.WriteLine($"Total number of pages: {pageCount}");
                    callflow.PageInfoFile.WriteLine($"Total number of pages: {pageCount}");

                    Page[] pageList = new Page[pageCount];
                    var multiPageGraph = new DirectedMultiGraph<VertexShape, EdgeShape>(); // used for multi-flow parsing (off-page references)

                    // pages start at page1 and go up to and including the page count
                    int i = 1;
                    foreach (var page in xPages)
                    {
                        var pageName = page?.Attribute("Name")?.Value;
                        Console.WriteLine($"Parsing page {i}: {pageName}");
                        callflow.PageInfoFile.WriteLine($"\n----Page {i}: {pageName}----");

                        XDocument xmlDoc = callflow.GetPageXML(i); // used for current xml being parsed

                        var graph = graphBuilder.BuildGraph(xmlDoc, i, pageName);
                        graphBuilder.MergePageGraph(graph, multiPageGraph); // combine all pages into one graph to parse between off-page references

                        pageList[i - 1] = BuildPage(graph, page, i);

                        // print out the graph data parsed from the page
                        callflow.PageInfoFile.WriteLine($"Page {i} has {graph.Vertices.Count} vertices and {graph.NumberOfEdges} edges");
                        //callflow.PageInfoFile.WriteLine($"\nGraph notation of page {i}:");
                        //PrintPageInformation(graph, callflow.PageInfoFile);

                        i++;
                    }

                    // once all of the pages have added their graphs together, connect each of their reference shapes together to join desired page flows
                    graphBuilder.ConnectReferenceShapes(multiPageGraph, pageList);

                    Console.WriteLine("Generating permutations...");
                    var pathSet = GetAllPermutations(multiPageGraph, callflow.PathOutputFile, callflow.Config, pageList, i);
                    var pathCount = pathSet.Count;

                    // find the minimum paths for test cases
                    if (pathCount > 0)
                    {
                        Console.WriteLine("Calculating minimum paths...");

                        Stopwatch stopwatch = new Stopwatch();
                        stopwatch.Start();

                        var minPathSet = GetMinimumPaths(multiPageGraph, pathSet, callflow.MinPathOutputFile);

                        stopwatch.Stop();
                        long elapsedMilliseconds = stopwatch.ElapsedMilliseconds;
                        Console.WriteLine($"Minimum path calculation time: {elapsedMilliseconds / 1000} seconds");

                        var minPathCount = minPathSet.Count;

                        callflow.MinPathOutputFile.WriteLine($"Minimum number of paths on Page {i} to cover all cases: {minPathCount}");
                        callflow.PageInfoFile.WriteLine($"Number of paths to test: {pathCount}\n");
                        pathCountTotal += pathCount;
                        minPathCountTotal += minPathCount;

                        Console.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
                        Console.WriteLine($"Total minimum number of paths to test: {minPathCountTotal}");
                        callflow.PathOutputFile.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
                        callflow.PathOutputFile.WriteLine($"Total minimum number of paths to test: {minPathCountTotal}");
                    }
                    else
                    {
                        Console.WriteLine("No paths found, please ensure the inputted configuration matches the actual IDs of the Visio");
                    }

                    callflow.ExecutionCleanup();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"An error occurred: {ex.Message}");

                    callflow.CleanupFiles();
                    Console.ReadLine(); // added for inspection of error
                }
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
                PrintPathID(allPaths, file);
            }
            else
            {
                file.WriteLine($"No test cases to be generated for this page");
            }

            return allPaths;
        }

        // method for finding all of the paths between every starting vertex to every ending vertex
        public static int i = 0;
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

        // find the minimum number of paths needed to cover every edge by picking the path that covers the most amount of uncovered edges at every iteration until all edges covered
        static List<List<VertexShape>> GetMinimumPaths(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<List<VertexShape>> allPaths, StreamWriter file)
        {
            // get unique edges from the graph and add them to the covered edges as the minimal paths become concrete
            var uniqueEdges = new HashSet<EdgeShape>(graph.Edges.Distinct());

            // set to keep track of covered edges
            var coveredEdges = new HashSet<EdgeShape>();

            // list of minimal paths covering all unique edges
            var minimalPathSet = new List<List<VertexShape>>();

            // while there are still uncovered edges
            while (coveredEdges.Count < uniqueEdges.Count)
            {
                // order the paths based on their amount of uncovered edges, then pick the path with the max amount of uncovered edges (i.e. the first path)
                var ordered = allPaths.Select(p => (uncovered: CountUncoveredEdges(graph, p, uniqueEdges, coveredEdges), path: p))
                    .ToList()
                    .OrderByDescending(p => p.uncovered);
                var max = ordered.Max(p => p.uncovered);
                var path = ordered.First().path;

                // if the max coverage path has edges then it should be included in our set of min paths
                if (path.Any() && max > 0)
                {
                    minimalPathSet.Add(path);
                    coveredEdges.UnionWith(GetEdgesFromPath(graph, path)); // update the set of covered edges to include the new path
                }
                else
                {
                    // otherwise there are no more paths which with provide more coverage
                    break;
                }
            }

            // print the minimal paths
            file.WriteLine("\nMinimal paths covering all unique edges:");
            PrintPathID(minimalPathSet, file);
            file.WriteLine("\nText for minimum paths:");
            PrintPathText(minimalPathSet, file);

            return minimalPathSet;
        }

        static int CountUncoveredEdges(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<VertexShape> path, HashSet<EdgeShape> uniqueEdges, HashSet<EdgeShape> coveredEdges)
            => GetEdgesFromPath(graph, path)
            .Intersect(uniqueEdges).Except(coveredEdges).Count();
        // Count(edge => uniqueEdges.Contains(edge) && !coveredEdges.Contains(edge));

        // find the edges between all pairs of consecutive vertices, then flatten the sequence and get just the unique edges to remove duplicates
        static IEnumerable<EdgeShape> GetEdgesFromPath(DirectedMultiGraph<VertexShape, EdgeShape> graph, IEnumerable<VertexShape> path)
            => path.Zip(path.Skip(1), graph.GetEdges).SelectMany(edges => edges).Distinct();


        static void PrintPathID(List<List<VertexShape>> pathSet, StreamWriter file)
        {
            int count = 1;
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
        }

        static void PrintPathText(List<List<VertexShape>> pathSet, StreamWriter file)
        {
            int count = 1;
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
        }
    }
}
