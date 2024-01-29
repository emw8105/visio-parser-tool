using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioParse.ConsoleHost
{
    /// <summary>
    /// A directional edge generated from pairs of connections, represented by a source and destination vertex
    /// </summary>
    public class EdgeShape
    {
        /// <summary>
        /// The identifier of the edge provided by the Visio
        /// </summary>
        public required string Id { get; init; }

        /// <summary>
        /// Text attached to the edge if any
        /// </summary>
        public string? Text { get; init; }

        /// <summary>
        /// ID referencing the destination vertex
        /// </summary>
        public required string ToShape { get; init; }

        /// <summary>
        /// ID referencing the source vertex
        /// </summary>
        public required string FromShape { get; init; }
    }

}
