using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace VisioParse.ConsoleHost
{
    public class VertexShape
    {
        public required string Id { get; init; }
        public string ?Text { get; set; }
        public string ?Type { get; init; }
        public required XElement XShape { get; init; }
        public required XElement XMaster { get; init; }
        public required string MasterId { get; init; }
    }
}
