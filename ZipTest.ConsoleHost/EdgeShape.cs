using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioParse.ConsoleHost
{
    public class EdgeShape
    {
        public required string Id { get; init; }
        public string Text { get; init; }
        public required string ToShape { get; init; } // string holding an ID referencing the destination vertex
        public required string FromShape { get; init; } // string holding an ID referencing the source vertex
    }

}
