using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;
using System.Linq;
using System.Collections.Immutable;


namespace VisioParse.ConsoleHost
{
    public class DirectedMultiGraph<TVertex, TEdge> : IGraph<TVertex, TEdge>
    {

        public DirectedMultiGraph() : this(
            outerMapFactory: () => new Dictionary<TVertex, IDictionary<TVertex, IList<TEdge>>>(),
            innerMapFactory: () => new Dictionary<TVertex, IList<TEdge>>(),
            listFactory: () => new List<TEdge>())
        {
            outgoingEdges = OuterMapFactory();
            incomingEdges = OuterMapFactory();
        }

        public DirectedMultiGraph(Func<IDictionary<TVertex, IDictionary<TVertex, IList<TEdge>>>> outerMapFactory,
            Func<IDictionary<TVertex, IList<TEdge>>> innerMapFactory, Func<IList<TEdge>> listFactory)
        {
            OuterMapFactory = outerMapFactory;
            InnerMapFactory = innerMapFactory;
            ListFactory = listFactory;
        }

        private Func<IDictionary<TVertex, IDictionary<TVertex, IList<TEdge>>>> OuterMapFactory { get; }

        private Func<IDictionary<TVertex, IList<TEdge>>> InnerMapFactory { get; }

        private Func<IList<TEdge>> ListFactory { get; }

        private readonly IDictionary<TVertex, IDictionary<TVertex, IList<TEdge>>> outgoingEdges;


        private readonly IDictionary<TVertex, IDictionary<TVertex, IList<TEdge>>> incomingEdges;

        private IDictionary<TVertex, IList<TEdge>> GetOutgoingEdgesMap(TVertex vertex)
        {
            if (outgoingEdges.TryGetValue(vertex, out var map))
                return map;
            map = InnerMapFactory();
            outgoingEdges[vertex] = map;
            incomingEdges[vertex] = InnerMapFactory();
            return map;
        }

        private IDictionary<TVertex, IList<TEdge>> GetIncomingEdgesMap(TVertex vertex)
        {
            if (incomingEdges.TryGetValue(vertex, out var map))
                return map;
            map = InnerMapFactory();
            outgoingEdges[vertex] = InnerMapFactory();
            incomingEdges[vertex] = map;
            return map;
        }

        public void Add(TVertex source, TVertex dest, TEdge data)
        {
            var outgoingMap = GetOutgoingEdgesMap(source);
            var incomingMap = GetIncomingEdgesMap(dest);

            if (!outgoingMap.TryGetValue(dest, out var outgoingList))
            {
                outgoingList = ListFactory();
                outgoingMap[dest] = outgoingList;
            }

            if (!incomingMap.TryGetValue(source, out var incomingList))
            {
                incomingList = ListFactory();
                incomingMap[source] = incomingList;
            }

            outgoingList.Add(data);
            incomingList.Add(data);
        }

        public bool AddVertex(TVertex vertex)
        {
            if (outgoingEdges.ContainsKey(vertex))
                return false;
            outgoingEdges[vertex] = InnerMapFactory();
            incomingEdges[vertex] = InnerMapFactory();
            return true;
        }

        public bool RemoveEdges(TVertex source, TVertex dest)
        {
            if (!outgoingEdges.ContainsKey(source))
                return false;
            if (!incomingEdges.ContainsKey(dest))
                return false;
            if (!outgoingEdges[source].ContainsKey(dest))
                return false;
            outgoingEdges[source].Remove(dest);
            incomingEdges[dest].Remove(source);
            return true;
        }

        public bool RemoveEdge(TVertex source, TVertex dest, TEdge data)
        {
            if (!outgoingEdges.ContainsKey(source))
                return false;
            if (!incomingEdges.ContainsKey(dest))
                return false;
            if (!outgoingEdges[source].ContainsKey(dest))
                return false;

            bool foundOut = outgoingEdges.ContainsKey(source) && outgoingEdges[source].ContainsKey(dest) && outgoingEdges[source][dest].Remove(data);
            bool foundIn = incomingEdges.ContainsKey(dest) && incomingEdges[dest].ContainsKey(source) && incomingEdges[dest][source].Remove(data);

            if (foundOut && !foundIn)
                throw new Exception("Edge found in outgoing but not incoming"); // TODO: Specialized Exception
            if (foundIn && !foundOut)
                throw new Exception("Edge found in incoming but not outgoing"); // TODO: Specialized Exception

            if (outgoingEdges.ContainsKey(source) && (!outgoingEdges[source].ContainsKey(dest) || outgoingEdges[source][dest].Count == 0))
                outgoingEdges[source].Remove(dest);
            if (incomingEdges.ContainsKey(dest) && (!incomingEdges[dest].ContainsKey(source) || incomingEdges[dest][source].Count == 0))
                incomingEdges[dest].Remove(source);

            return foundOut;
        }

        public bool RemoveVertex(TVertex vertex)
        {
            if (!outgoingEdges.ContainsKey(vertex))
                return false;
            foreach (var other in outgoingEdges[vertex].Keys)
                incomingEdges[other].Remove(vertex);
            foreach (var other in incomingEdges[vertex].Keys)
                outgoingEdges[other].Remove(vertex);
            outgoingEdges.Remove(vertex);
            incomingEdges.Remove(vertex);
            return true;
        }

        public bool RemoveVertices(IEnumerable<TVertex> vertices)
        {
            bool changed = false;
            foreach (var vertex in vertices)
            {
                if (RemoveVertex(vertex))
                    changed = true;
            }
            return changed;
        }

        public int NumberOfVertices => outgoingEdges.Count;
        public int NumberOfEdges => outgoingEdges.Values.Sum(outer => outer.Values.Sum(inner => inner.Count));
        //=> CalculateNumberOfEdges();

        private int CalculateNumberOfEdges()
        {
            int count = 0;
            foreach (var sourceEntry in outgoingEdges.Values)
                foreach (var destEntry in sourceEntry.Values)
                    count += destEntry.Count;
            return count;
        }

        public IEnumerable<TEdge> GetOutgoingEdges(TVertex vertex)
        {
            if (!outgoingEdges.ContainsKey(vertex))
                return Enumerable.Empty<TEdge>();
            return incomingEdges[vertex].SelectMany(v => v.Value);
        }

        public IEnumerable<TEdge> GetIncomingEdges(TVertex vertex)
        {
            if (!incomingEdges.ContainsKey(vertex))
                return Enumerable.Empty<TEdge>();
            return incomingEdges[vertex].SelectMany(v => v.Value);
        }

        public ISet<TVertex> GetParents(TVertex vertex)
            => (incomingEdges.TryGetValue(vertex, out var parentMap)
                ? parentMap.Keys
                : Enumerable.Empty<TVertex>())
            .ToImmutableHashSet();

        public ISet<TVertex> GetChildren(TVertex vertex)
            => (outgoingEdges.TryGetValue(vertex, out var childMap)
                ? childMap.Keys
                : Enumerable.Empty<TVertex>())
            .ToImmutableHashSet();

        public ISet<TVertex> GetNeighbors(TVertex vertex)
        {
            var neighbors = new List<TVertex>(GetChildren(vertex));
            neighbors.AddRange(GetParents(vertex));
            return neighbors.ToImmutableHashSet();
        }

        public void Clear()
        {
            incomingEdges.Clear();
            outgoingEdges.Clear();
        }

        public bool ContainsVertex(TVertex vertex) => outgoingEdges.ContainsKey(vertex);

        public bool IsEdge(TVertex source, TVertex dest)
        {
            if (!outgoingEdges.TryGetValue(source, out var childrenMap) || !childrenMap.Any())
                return false;
            if (!childrenMap.TryGetValue(dest, out var edges))
                return false;
            return edges.Any();
        }

        public bool IsNeighbor(TVertex source, TVertex dest) => IsEdge(source, dest) || IsEdge(dest, source);

        public ISet<TVertex> Vertices => outgoingEdges.Keys.ToImmutableHashSet();

        public IEnumerable<TEdge> Edges => outgoingEdges.Values.SelectMany(inner => inner.Values).SelectMany(edges => edges);

        public bool IsEmpty() => !outgoingEdges.Any();

        public void RemoveZeroDegreeNodes()
        {
            var toDelete = outgoingEdges.Keys.Where(vertex => !outgoingEdges[vertex].Any() && !incomingEdges[vertex].Any());
            foreach (var vertex in toDelete)
            {
                outgoingEdges.Remove(vertex);
                incomingEdges.Remove(vertex);
            }
        }

        public IEnumerable<TEdge> GetEdges(TVertex source, TVertex dest)
        {
            if (!outgoingEdges.TryGetValue(source, out var childrenMap))
                return Enumerable.Empty<TEdge>();
            if (!childrenMap.TryGetValue(dest, out var edges))
                return Enumerable.Empty<TEdge>();
            return edges;
        }

        //public IEnumerable<TVertex> GetShortestPath(TVertex node1, TVertex node2)
        //{

        //}

        //public IEnumerable<TVertex> GetShortestPath(TVertex node1, TVertex node2, bool directionSensitive)
        //{

        //}

        public int GetInDegree(TVertex vertex)
        {
            if (incomingEdges.TryGetValue(vertex, out var map))
                return map.Values.Sum(inner => inner.Count);
            return 0;
        }

        public int GetOutDegree(TVertex vertex)
        {
            if (outgoingEdges.TryGetValue(vertex, out var map))
                return map.Values.Sum(inner => inner.Count);
            return 0;
        }

        public IEnumerable<ISet<TVertex>> ConnectedComponents { get; }

        public void GetObjectData(SerializationInfo info, StreamingContext context) => throw new NotImplementedException();
    }
}
