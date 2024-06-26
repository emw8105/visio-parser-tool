﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioParse.ConsoleHost
{
    public class GraphBuilder
    {
        private CallflowHandler Callflow { get; set; }

        private static readonly XNamespace ns = "http://schemas.microsoft.com/office/visio/2012/main"; // needed for queries to match elements, use for all xml parsing

        public GraphBuilder(CallflowHandler callflow)
        {
            Callflow = callflow;
        }

        public DirectedMultiGraph<VertexShape, EdgeShape> BuildGraph(XDocument xmlDoc, int pageNum, string pageName)
        {
            // shapes include both vertices and edges, some pages have separate shapes within shapes that are missing various properties, must handle if using this
            // var shapes = xmlDoc.Root.Element(ns + "Shapes").Elements(ns + "Shape"); // use this parsing instead to only get top-level shapes
            var shapes = xmlDoc.Descendants(ns + "Shape");
            var connections = xmlDoc.Descendants(ns + "Connect");

            DirectedMultiGraph<VertexShape, EdgeShape> graph = new DirectedMultiGraph<VertexShape, EdgeShape>(); // different graph for each page

            // need to separate the shapes into different categories for graph functionality
            // need to compute edges first because connections are recorded as shapes, don't want to add an edge as a vertex
            var connectionShapes = new HashSet<string>();
            List<EdgeShape> pageEdges = new List<EdgeShape>();

            // edges are stored as a pair of connections, must parse both to find the origin node and the destination node
            // once a pair of connections is found, we store it as an edge
            MatchConnections(connections, connectionShapes, pageEdges, shapes);

            // Extract and print shape information from the current page, convert each non-edge shape into a vertex
            ExtractShapeToVertex(graph, shapes, connectionShapes, Callflow.PageInfoFile, pageName);

            // when shapes are converted to vertices, the text is edited so save the edited page
            xmlDoc.Save(Callflow.ExtractPath + @"\visio\pages\page" + pageNum + ".xml");

            // unfortunately, can't add an edge without the vertex existing
            // but can't add vertexes until we determine which shapes are connections
            // this is used to assign the edge placements in the graph themselves using their stored data
            AssignEdges(pageEdges, graph);

            return graph;
        }

        private static void MatchConnections(IEnumerable<XElement> connections, HashSet<string> connectionShapes, List<EdgeShape> pageEdges, IEnumerable<XElement> shapes)
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

                    var edgeShape = shapes.First(shape => shape?.Attribute("ID")?.Value == fromSheetBegin); // get the text of the edge from it's respective shape

                    EdgeShape edge = new()
                    {
                        Id = fromSheetBegin,
                        ToShape = toSheetEnd,
                        FromShape = toSheetBegin,
                        Text = edgeShape?.Element(ns + "Text")?.Value
                    };
                    pageEdges.Add(edge);

                    connectionShapes.Add(fromSheetBegin); // add the shape ID to a list of shapes designated as connections
                    endXConnections.Remove(fromSheetBegin); // remove the corresponding connection's shape ID
                }
                else
                {
                    Console.WriteLine($"Error: Found 'BeginX' without corresponding 'EndX'. Sheet: {fromSheetBegin}, note that only directional arrows are parsed");
                }
            }

            // if any shapes are still leftover in the endXConnections, then not every connection was matched, throw an error
            foreach (var endXConnection in endXConnections.Values)
            {
                string fromSheetEnd = endXConnection.Attribute("FromSheet").Value;

                Console.WriteLine($"Error: Found 'EndX' without corresponding 'BeginX'. Sheet: {fromSheetEnd}, note that only directional arrows are parsed");
            }
        }

        private static void ExtractShapeToVertex(DirectedMultiGraph<VertexShape, EdgeShape> graph, IEnumerable<XElement>? shapes, HashSet<string> connectionShapes, StreamWriter file, string pageName)
        {
            // attributes to parse from the shapes
            string id;
            string? type;
            string? masterId;

            // Console.WriteLine(String.Join(",", connectionShapes)); // checks if connections were assigned properly
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
                                PageName = pageName
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

                            // first see if the shape has a hyperlink section, if so then drill down to get the subaddress which contains the page name for the off-page reference
                            var hyperlinkSection = shape.Descendants(ns + "Section").FirstOrDefault(section => section.Attribute("N")?.Value == "Hyperlink");

                            if (hyperlinkSection != null)
                            {
                                var subAddress = hyperlinkSection.Descendants(ns + "Row").Descendants(ns + "Cell").FirstOrDefault(cell => cell.Attribute("N")?.Value == "SubAddress")?.Attribute("V")?.Value;
                                vertex.PageReference = subAddress;
                                Console.WriteLine($"Shape {vertex.Id} has an off-page reference to: {subAddress}");
                                file.WriteLine($"Shape {vertex.Id} has an off-page reference to: {subAddress}");
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

                        // after shape information is saved into corresponding object, write the ID to the text of the shape to visualize permutations
                        WriteShapeID(shape);
                    }
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine($"Null reference exception encountered:");
                    Console.WriteLine(ex);
                }
            }
        }

        private static void AssignEdges(List<EdgeShape> pageEdges, DirectedMultiGraph<VertexShape, EdgeShape> graph)
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

        private static void WriteShapeID(XElement shape)
        {
            var shapeId = shape?.Attribute("ID")?.Value;
            var masterId = shape?.Attribute("Master")?.Value;
            var newText = $"ID: {shapeId}";

            // check if the text element exists and update its value
            var textElement = shape?.Element(ns + "Text");
            if (textElement is not null)
            {
                if (masterId is not null)
                {
                    textElement.Value = newText + $"\nMasterID: {masterId}";
                }
                else
                {
                    textElement.Value = newText;
                }
            }
            else
            {
                // if text element does not exist, create and add it
                shape.Add(new XElement(ns + "Text", newText));
            }
        }

        public void MergePageGraph(DirectedMultiGraph<VertexShape, EdgeShape> graph, DirectedMultiGraph<VertexShape, EdgeShape> mergedGraph)
        {
            // add every vertex
            graph.Vertices.ToList().ForEach(vertex => mergedGraph.AddVertex(vertex));

            // add every edge
            graph.Edges.ToList().ForEach(edge =>
            {
                mergedGraph.Add(
                    graph.Vertices.First(v => v.Id == edge.FromShape),
                    graph.Vertices.First(v => v.Id == edge.ToShape),
                    edge
                );
            });
        }

        // On-Page References are the same as checkpoints
        public void ConnectReferenceShapes(DirectedMultiGraph<VertexShape, EdgeShape> graph, Page[] pageList)
        {
            var nodeOption = Callflow.Config.NodeOption;
            var checkpointContent = Callflow.Config.CheckpointContent;

            // Note: for references, end nodes should point to start nodes to join the pages together
            HashSet<string> pageNames = new HashSet<string>(pageList.Select(p => p.Name));
            switch (nodeOption.Substring(1, 1))
            {
                case "1":
                    // off-page references
                    var referenceStartNodes = graph.Vertices.Where(vertex => !vertex.PageReference.Equals("") && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    var referenceEndNodes = graph.Vertices.Where(vertex => !vertex.PageReference.Equals("") && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);

                    // match references together and generate an edge to link them
                    CreateReferenceEdges(referenceStartNodes, referenceEndNodes, graph);
                    break;
                case "2":
                    // checkpoints (format of text needs to be "pageName: identifier")
                    var checkpointStartNodes = graph.Vertices.Where(vertex => vertex.MasterId == checkpointContent && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    var checkpointEndNodes = graph.Vertices.Where(vertex => vertex.MasterId == checkpointContent && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);
                    CreateCheckpointEdges(checkpointStartNodes, checkpointEndNodes, graph);

                    break;
                case "3":
                    // both
                    referenceStartNodes = graph.Vertices.Where(vertex => !vertex.PageReference.Equals("") && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    referenceEndNodes = graph.Vertices.Where(vertex => !vertex.PageReference.Equals("") && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);
                    checkpointStartNodes = graph.Vertices.Where(vertex => vertex.MasterId == checkpointContent && graph.GetInDegree(vertex) == 0 && graph.GetOutDegree(vertex) > 0);
                    checkpointEndNodes = graph.Vertices.Where(vertex => vertex.MasterId == checkpointContent && graph.GetOutDegree(vertex) == 0 && graph.GetInDegree(vertex) > 0);

                    CreateReferenceEdges(referenceStartNodes, referenceEndNodes, graph);
                    CreateCheckpointEdges(checkpointStartNodes, checkpointEndNodes, graph);
                    break;
                default:
                    // none
                    break;
            }
        }

        private void CreateCheckpointEdges(IEnumerable<VertexShape>? checkpointStartNodes, IEnumerable<VertexShape>? checkpointEndNodes, DirectedMultiGraph<VertexShape, EdgeShape> graph)
        {
            //Console.Write("Start node on-page references: ");
            //checkpointStartNodes.ToList().ForEach(n => Console.Write(n.Text + ", "));
            //Console.WriteLine();
            //Console.Write("End node on-page references: ");
            //checkpointEndNodes.ToList().ForEach(n => Console.Write(n.Text + ", "));
            //Console.WriteLine();

            // on-page reference connecting
            Console.WriteLine("Checking for on-page references...");
            Dictionary<string, VertexShape> startNodesMap = new Dictionary<string, VertexShape>();

            // populate the dictionaries with nodes
            foreach (var startNode in checkpointStartNodes)
            {
                string key = $"{startNode.PageName}_{startNode.Text}";
                startNodesMap[key] = startNode;
            }

            // check for connections using the dictionaries
            foreach (var endNode in checkpointEndNodes)
            {
                var endNodeKey = $"{endNode.PageName}_{endNode.Text}";
                if (startNodesMap.ContainsKey(endNodeKey))
                {
                    var startNode = startNodesMap[endNodeKey];

                    Console.WriteLine($"CREATING EDGE BETWEEN ON-PAGE REFERENCES: {endNode.Id + " (" + endNode.Text + ")"} and {startNode.Id + " (" + startNode.Text + ") "}");

                    var referenceEdge = new EdgeShape
                    {
                        Id = Guid.NewGuid().ToString(),
                        Text = "Reference link",
                        ToShape = endNode.Id,
                        FromShape = startNode.Id
                    };

                    graph.Add(endNode, startNode, referenceEdge);
                }
                else
                {
                    // detect unmatched connections
                    Console.WriteLine($"No matching start node found for end node: {endNodeKey}");
                }
            }
        }

        private void CreateReferenceEdges(IEnumerable<VertexShape>? referenceStartNodes, IEnumerable<VertexShape>? referenceEndNodes, DirectedMultiGraph<VertexShape, EdgeShape> graph)
        {
            Console.WriteLine("Checking for off-page references...");
            //Console.Write("Start node off-page references: ");
            //referenceStartNodes.ToList().ForEach(n => Console.Write(n.Text + ", "));
            //Console.WriteLine();
            //Console.Write("End node off-page references: ");
            //referenceEndNodes.ToList().ForEach(n => Console.Write(n.Text + ", "));
            //Console.WriteLine();

            // match references together and generate an edge to link them
            Dictionary<string, VertexShape> startNodesMap = new Dictionary<string, VertexShape>();

            // populate the dictionaries with nodes
            foreach (var startNode in referenceStartNodes)
            {
                string key = $"{startNode.PageName}_{startNode.PageReference}";
                startNodesMap[key] = startNode;
            }

            foreach (var endNode in referenceEndNodes)
            {
                var endNodeKey = $"{endNode.PageReference}_{endNode.PageName}";
                if (startNodesMap.ContainsKey(endNodeKey))
                {
                    var startNode = startNodesMap[endNodeKey];

                    Console.WriteLine($"CREATING EDGE BETWEEN OFF-PAGE REFERENCES: FROM {startNode.PageName} TO {endNode.PageName} (From {startNode.Id} to {endNode.Id})");

                    var referenceEdge = new EdgeShape
                    {
                        Id = Guid.NewGuid().ToString(),
                        Text = "Reference link",
                        ToShape = endNode.Id,
                        FromShape = startNode.Id
                    };

                    graph.Add(endNode, startNode, referenceEdge);
                }
                else
                {
                    // detect unmatched connections
                    Console.WriteLine($"No matching start node found for end node: {endNodeKey}");

                    // if there are no found matching references, check if there is a singular start node on the referenced page to create an edge to
                    var startNodes = graph.Vertices.Where(vertex => vertex.PageName == endNode.PageReference && vertex.MasterId == Callflow.Config.StartNodeContent && graph.GetInDegree(vertex) > 0 && graph.GetOutDegree(vertex) == 0);
                    if (startNodes.Count() == 1)
                    {
                        var startNode = startNodes.First();
                        var referenceEdge = new EdgeShape
                        {
                            Id = Guid.NewGuid().ToString(),
                            Text = "Reference link",
                            ToShape = endNode.Id,
                            FromShape = startNode.Id
                        };

                        graph.Add(endNode, startNode, referenceEdge);
                        Console.WriteLine($"Reference {endNode.PageReference}: {endNodeKey}");
                    }
                }
            }
        }
    }
}
