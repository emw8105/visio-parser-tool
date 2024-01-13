﻿using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioParse.ConsoleHost
{
    /// <summary>
    /// A vertex generated from non-connection shapes storing information about the node
    /// </summary>
    public class VertexShape
    {
        /// <summary>
        /// The identifier of the vertex provided by the Visio
        /// </summary>
        public required string Id { get; init; }

        /// <summary>
        /// The text attached to the shape if any
        /// </summary>
        public string? Text { get; set; } = string.Empty;

        /// <summary>
        /// The type of a shape if any, most will just be 'Shape'
        /// </summary>
        public string? Type { get; init; } = string.Empty;

        /// <summary>
        /// The XML content of the shape
        /// </summary>
        public required XElement XShape { get; init; }

        /// <summary>
        /// the XML content of the master shape (could be doing this wrong)
        /// </summary>
        public required XElement XMaster { get; init; }

        /// <summary>
        /// 
        /// </summary>
        public required string PageName = string.Empty;

        /// <summary>
        /// The identifier representing the master of the shape
        /// </summary>
        public required string MasterId { get; init; }

        /// <summary>
        /// The page name serving as a reference if the vertex is an off-page reference
        /// </summary>
        public string? pageReference = string.Empty;

        /// <summary>
        /// The specific vertex Id serving as a reference if the vertex is an off-page reference to a specific node
        /// </summary>
        public string? vertexReference = string.Empty;
    }
}
