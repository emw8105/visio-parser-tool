using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace VisioParse.ConsoleHost
{
    /// <summary>
    /// Represents a page in the visio file containing its shapes as a graph and identifying information
    /// </summary>
    public class Page
    {
        /// <summary>
        /// Stores the page's parsed vertices and edges
        /// </summary>
        public DirectedMultiGraph<VertexShape, EdgeShape>? Graph { get; set; }

        /// <summary>
        /// Internal Id stored within the Visio file
        /// </summary>
        public required string Id { get; set; }

        /// <summary>
        /// Linear position used for XML file name and array positioning for multi-flow paths
        /// </summary>
        public required int Number { get; set; }

        /// <summary>
        /// The name designated by the user for the page
        /// </summary>
        public required string Name { get; set; }

    }
}
