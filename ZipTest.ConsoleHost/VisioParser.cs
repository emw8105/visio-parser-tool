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
using System.IO.MemoryMappedFiles;
using ArchitectConvert.ConsoleHost;
using System.Drawing;

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
            callflow.ConfigurationSetup();

            Console.WriteLine("extracting file to " + callflow.ExtractPath);
            ZipFile.ExtractToDirectory(callflow.ZipPath, callflow.ExtractPath); // convert given visio file to xml components
            Console.WriteLine("finished extraction, parsing components...");

            // read extracted xml file contents
            try
            {
                // first parse "pages.xml" to find page count
                XDocument pagesXml;
                using (XmlTextReader documentReader = new XmlTextReader(callflow.ExtractPath + @"\visio\pages\pages.xml"))
                {
                    pagesXml = XDocument.Load(documentReader);
                }

                var pages = pagesXml.Descendants(ns + "Page");
                int pageCount = pages.Count();
                int pathCountTotal = 0;

                Console.WriteLine($"Total number of pages: {pageCount}");
                callflow.PageInfoFile.WriteLine($"Total number of pages: {pageCount}");

                XDocument xmlDoc; // used for current xml being parsed

                DirectedMultiGraph<VertexShape, EdgeShape>[] pageGraphs = new DirectedMultiGraph<VertexShape, EdgeShape>[pageCount];

                // pages start at page1 and go up to and including the page count
                for (int i = 1; i <= pageCount; i++)
                {
                    Console.WriteLine($"Parsing page {i}");
                    callflow.PageInfoFile.WriteLine($"\n----Page {i}----");

                    using (XmlTextReader reader = new XmlTextReader(callflow.ExtractPath + @"\visio\pages\page" + i + ".xml"))
                    {
                        xmlDoc = XDocument.Load(reader);
                    }

                    // shapes include both vertices and edges, some pages have separate shapes within shapes that are missing various properties, must handle if using this
                    // var shapes = xmlDoc.Root.Element(ns + "Shapes").Elements(ns + "Shape"); // use this parsing instead to only get top-level shapes
                    var shapes = xmlDoc.Descendants(ns + "Shape");
                    var connections = xmlDoc.Descendants(ns + "Connect");

                    DirectedMultiGraph<VertexShape, EdgeShape> graph = new DirectedMultiGraph<VertexShape, EdgeShape>(); // different graph for each page

                    // need to separate the shapes into different categories for graph functionality
                    // need to compute edges first because connections are recorded as shapes, don't want to add an edge as a vertex
                    var connectionShapes = new HashSet<string>();
                    List<VertexShape> pageShapes = new List<VertexShape>();
                    List<EdgeShape> pageEdges = new List<EdgeShape>();

                    // edges are stored as a pair of connections, must parse both to find the origin node and the destination node
                    // once a pair of connections is found, we store it as an edge
                    MatchConnections(connections, connectionShapes, pageEdges);

                    // Extract and print shape information from the current page, convert each non-edge shape into a vertex
                    ExtractShapeToVertex(graph, shapes, connectionShapes, callflow.PageInfoFile);

                    // when shapes are converted to vertices, the text is edited so save the edited page
                    xmlDoc.Save(callflow.ExtractPath + @"\visio\pages\page" + i + ".xml");

                    // unfortunately, can't add an edge without the vertex existing
                    // but can't add vertexes until we determine which shapes are connections
                    // this is used to assign the edge placements in the graph themselves using their stored data
                    AssignEdges(pageEdges, graph);

                    // sometimes tables and other extra visual elements are added as shapes, remove them to reduce clutter and save the graph
                    graph.RemoveZeroDegreeNodes();
                    pageGraphs[i-1] = graph;   

                    // print out the graph data parsed from the page
                    callflow.PageInfoFile.WriteLine($"Page {i} has {graph.Vertices.Count} vertices and {graph.NumberOfEdges} edges");
                    callflow.PageInfoFile.WriteLine($"\nGraph notation of page {i}:");
                    PrintPageInformation(graph, callflow.PageInfoFile);

                    // find the permutations, every path from every starting node to every ending node, each path is a test case
                    callflow.PathOutputFile.WriteLine($"\nPaths of page {i}:");
                    int numPaths = GetAllPermutations(graph, callflow.PathOutputFile, callflow.NodeOption, callflow.StartNodeContent, callflow.EndNodeContent, i);
                    callflow.PageInfoFile.WriteLine($"Number of paths to test: {numPaths}");
                    Console.WriteLine($"Number of paths to test: {numPaths}");
                    pathCountTotal += numPaths;
                }
                Console.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
                callflow.PathOutputFile.WriteLine($"\nTotal number of paths to test to cover every page: {pathCountTotal}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }


            callflow.PageInfoFile.Flush();
            callflow.PageInfoFile.Close();
            callflow.PathOutputFile.Flush();
            callflow.PathOutputFile.Close();

            callflow.ExecutionCleanup();
        }

        static void GetSpecifiedNodes(DirectedMultiGraph<VertexShape, EdgeShape> graph, string nodeOption, string? startNodeContent, string? endNodeContent, out IEnumerable<VertexShape> startNodes, out IEnumerable<VertexShape> endNodes)
        {
            switch (nodeOption)
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

        static int GetAllPermutations(DirectedMultiGraph<VertexShape, EdgeShape> graph, StreamWriter file, string nodeOption, string? startNodeContent, string? endNodeContent, int pageNum)
        {
            // first get the start and end nodes based on user specification
            IEnumerable<VertexShape>? startNodes;
            IEnumerable<VertexShape>? endNodes;
            GetSpecifiedNodes(graph, nodeOption, startNodeContent, endNodeContent, out startNodes, out endNodes);
            file.Write("Start nodes: ");
            startNodes.ToList().ForEach(node => file.Write($"{node.Id}, "));
            file.Write("\nEnd nodes: ");
            endNodes.ToList().ForEach(node => file.Write($"{node.Id}, "));
            file.WriteLine();

            Console.Write("Start nodes: ");
            startNodes.ToList().ForEach(node => Console.Write($"{node.Id}, "));
            Console.Write("\nEnd nodes: ");
            endNodes.ToList().ForEach(node => Console.Write($"{node.Id}, "));
            Console.WriteLine();

            // list containing the paths containing each node
            var allPaths = new List<List<VertexShape>>();

            //file.WriteLine("\nPaths:");
            int numPaths = 0;
            int count = 1;

            // find every path from every start node to every end node
            foreach (var startNode in startNodes)
            {
                foreach (var endNode in endNodes)
                {
                    var permutations = FindPermutations(graph, startNode, endNode);
                    allPaths.AddRange(permutations);
                    numPaths += permutations.Count;

                    foreach (var currentPath in permutations)
                    {
                        file.Write($"{count}. ");
                        foreach (var node in currentPath)
                        {
                            file.Write($"{node.Id} -> ");
                        }
                        file.WriteLine();
                        count++;
                    }
                }
            }

            // print the text for every path, not very necessary
            //count = 1;
            //file.WriteLine("\nText for above paths:");
            //foreach (var path in allPaths)
            //{
            //    file.Write($"{count}. ");
            //    foreach (var node in path)
            //    {
            //        if (node.Text != null)
            //        {
            //            file.Write($"{node.Text.Trim('\0')} -> ");
            //        }
            //    }
            //    file.WriteLine();
            //    count++;
            //}
            var minimalPathSet = GetMinimumPaths(graph, allPaths, file);
            file.WriteLine($"Total number of minimal paths on page {pageNum} to cover all edges: {minimalPathSet.Count}");

            file.WriteLine($"Total number of paths on page {pageNum} to test: {numPaths}");

            return numPaths;
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

        static List<List<VertexShape>> GetMinimumPaths(DirectedMultiGraph<VertexShape, EdgeShape> graph, List<List<VertexShape>> allPaths, StreamWriter file)
        {
            // get unique edges from the graph and add them to the covered edges as the minimal paths become concrete
            var uniqueEdges = graph.Edges.Distinct().ToList();
            HashSet<EdgeShape> coveredEdges = new HashSet<EdgeShape>();

            // list of minimal paths covering all unique edges
            var minimalPathSet = new List<List<VertexShape>>();

            // loop through paths and add to minimalPathSet if it covers a unique edge
            foreach (var path in allPaths)
            {
                var pathCoversUniqueEdge = false;

                for (int i = 0; i < path.Count - 1; i++)
                {
                    VertexShape source = path[i];
                    VertexShape target = path[i + 1];

                    // find the edge from source to target
                    if (graph.GetEdges(source, target).Any(edge => uniqueEdges.Contains(edge) && !coveredEdges.Contains(edge)))
                    {
                        pathCoversUniqueEdge = true;
                        coveredEdges.UnionWith(graph.GetEdges(source, target));
                    }
                }

                if (pathCoversUniqueEdge)
                {
                    minimalPathSet.Add(path);
                }
            }

            // print the minimal paths
            int count = 1;
            file.WriteLine("\nMinimal paths covering all unique edges:");
            foreach (var path in minimalPathSet)
            {
                file.Write($"{count}. ");
                foreach (var node in path)
                {
                    file.Write($"{node.Id} -> ");
                }
                file.WriteLine();
                count++;
            }
            count = 1;
            file.WriteLine("\nText for above paths:");
            foreach (var path in allPaths)
            {
                file.Write($"{count}. ");
                foreach (var node in path)
                {
                    if (node.Text != null)
                    {
                        file.Write($"{node.Text.Trim('\0')} -> ");
                    }
                }
                file.WriteLine();
                count++;
            }

            return minimalPathSet;
        }

        static void MatchConnections(IEnumerable<XElement>? connections, HashSet<string> connectionShapes, List<EdgeShape> pageEdges)
        {
            // create two dictionaries, load all connections into each, then match them together based on shapeID or their "FromSheet" value to create an edge
            var beginXConnections = new Dictionary<string, XElement>();
            var endXConnections = new Dictionary<string, XElement>();

            foreach (var connect in connections)
            {
                string fromSheet = connect.Attribute("FromSheet").Value; // the Shape ID of the edge
                string fromCell = connect.Attribute("FromCell").Value; // the direction of the edge (begin or end)

                // put origin node portion of edges into beginXConnections and destination node portion of edges into endXConnections
                if (fromCell == "BeginX")
                {
                    beginXConnections[fromSheet] = connect;
                }
                else if (fromCell == "EndX")
                {
                    endXConnections[fromSheet] = connect;
                }
            }

            // loop through the BeginX connections to match with the corresponding EndX connections and create an edge
            foreach (var beginXConnection in beginXConnections.Values)
            {
                string fromSheetBegin = beginXConnection.Attribute("FromSheet").Value; // Edge shape ID
                string toSheetBegin = beginXConnection.Attribute("ToSheet").Value; // origin node of the edge

                if (endXConnections.TryGetValue(fromSheetBegin, out var endXConnection))
                {
                    string toSheetEnd = endXConnection.Attribute("ToSheet").Value; // destination node of the edge

                    EdgeShape edge = new EdgeShape()
                    {
                        Id = fromSheetBegin,
                        ToShape = toSheetEnd,
                        FromShape = toSheetBegin,
                    };
                    pageEdges.Add(edge);

                    string edgeId = $"EdgeID:{fromSheetBegin}_from:{toSheetBegin}_to:{toSheetEnd}";

                    // check if the edge is already processed
                    if (!connectionShapes.Contains(edgeId))
                    {
                        connectionShapes.Add(fromSheetBegin); // add the shape ID to a list of shapes designated as connections
                    }

                    endXConnections.Remove(fromSheetBegin); // remove the corresponding connection's shape ID
                }
                else
                {
                    Console.WriteLine($"Error: Found 'BeginX' without corresponding 'EndX'. Sheet: {fromSheetBegin}, please use single-directional arrows only");
                }
            }

            // if any shapes are still leftover in the endXConnections, then not every connection was matched, throw an error
            foreach (var endXConnection in endXConnections.Values)
            {
                string fromSheetEnd = endXConnection.Attribute("FromSheet").Value;

                Console.WriteLine($"Error: Found 'EndX' without corresponding 'BeginX'. Sheet: {fromSheetEnd}, please use single-directional arrows only");
            }
        }

        static void ExtractShapeToVertex(DirectedMultiGraph<VertexShape, EdgeShape> graph, IEnumerable<XElement>? shapes, HashSet<string> connectionShapes, StreamWriter file)
        {
            // attributes to parse from the shapes
            string id;
            string? type;
            string? masterId;

            foreach (var shape in shapes)
            {
                try
                {
                    // first check if the given shape is actually a connection and not a vertex
                    id = shape.Attribute("ID").Value;
                    if (!connectionShapes.Contains(id))
                    {
                        var typeAttribute = shape.Attribute("Type"); // type should only be null if we have a subshape
                        masterId = shape?.Attribute("Master")?.Value; // some shapes not associated with a specific master

                        if (typeAttribute != null)
                        {
                            VertexShape vertex = new()
                            {
                                Id = id,
                                Type = typeAttribute.Value,
                                XMaster = shape.Element("Master"),
                                MasterId = masterId,
                                XShape = shape,
                            };

                            type = typeAttribute.Value;
                            var textElement = shape.Element(ns + "Text"); // not all shapes have text, so we need a null check
                            if (textElement != null)
                            {
                                string text = textElement.Value.Trim();
                                vertex.Text = text;
                                file.WriteLine($"Shape ID: {id}, Type: {type}, Master = {(masterId != null ? masterId : "null")}, Text: {text}");

                            }
                            else
                            {
                                file.WriteLine($"Shape ID: {id}, Type: {type}, Master = {(masterId != null ? masterId : "null")}");
                            }

                            graph.AddVertex(vertex);
                        }
                        else
                        {
                            // subshapes have a property called "MasterShape" instead of Master
                            masterId = shape?.Attribute("MasterShape")?.Value;
                            file.WriteLine($"ShapeID: {id} is a subshape, MasterShape = {masterId}");
                            Console.WriteLine($"Subshape detected, it will not be included in the path determination - ShapeID: {id}");
                        }
                    }

                    // after shape information is saved into corresponding object, write the ID to the text of the shape to visualize permutations
                    WriteShapeID(shape);
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine($"Null reference exception encountered:");
                    Console.WriteLine(ex);
                }
            }
        }

        static void AssignEdges(List<EdgeShape> pageEdges, DirectedMultiGraph<VertexShape, EdgeShape> graph)
        {
            foreach (var edge in pageEdges)
            {
                // Find the fromShape and toShape based on their IDs
                var fromShape = graph.Vertices.FirstOrDefault(shape => shape.Id == edge.FromShape);
                var toShape = graph.Vertices.FirstOrDefault(shape => shape.Id == edge.ToShape);

                if (fromShape != null && toShape != null)
                {
                    graph.Add(fromShape, toShape, edge);
                }
                else
                {
                    Console.WriteLine($"Error: Couldn't find source vertex {edge.FromShape} or destination vertex {edge.ToShape} for edge {edge.Id}");
                }
            }
        }

        static void WriteShapeID(XElement shape)
        {
            string shapeId = shape.Attribute("ID").Value;
            string newText = $"ID: {shapeId}";

            // check if the text element exists and update its value
            var textElement = shape.Element(ns + "Text");
            if (textElement != null)
            {
                textElement.Value = newText;
            }
            else
            {
                // if text element does not exist, create and add it
                shape.Add(new XElement(ns + "Text", newText));
            }
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
            //foreach (var vertex in graph.Vertices)
            //{
            //    file.Write($"Neighbors of Vertex: {vertex.Id} - ");
            //    var neighbors = graph.GetNeighbors(vertex);
            //    foreach (var neighbor in neighbors)
            //    {
            //        file.Write(neighbor.Id + ", ");
            //    }
            //    file.WriteLine();
            //}
        }
    }
}