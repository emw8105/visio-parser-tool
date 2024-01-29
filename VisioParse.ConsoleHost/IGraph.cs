using System;
using System.Collections.Generic;
using System.Runtime.Serialization;
using System.Text;

namespace VisioParse.ConsoleHost
{
    public interface IGraph<TVertex, TEdge> : ISerializable
    {
        void Add(TVertex source, TVertex dest, TEdge data);

        bool AddVertex(TVertex vertex);

        bool RemoveEdges(TVertex source, TVertex dest);

        bool RemoveEdge(TVertex source, TVertex dest, TEdge data);

        bool RemoveVertex(TVertex vertex);

        bool RemoveVertices(IEnumerable<TVertex> vertices);

        int NumberOfVertices { get; }

        int NumberOfEdges { get; }

        IEnumerable<TEdge> GetOutgoingEdges(TVertex vertex);

        IEnumerable<TEdge> GetIncomingEdges(TVertex vertex);

        ISet<TVertex> GetParents(TVertex vertex);

        ISet<TVertex> GetChildren(TVertex vertex);

        ISet<TVertex> GetNeighbors(TVertex vertex);

        void Clear();

        bool ContainsVertex(TVertex vertex);

        bool IsEdge(TVertex source, TVertex dest);

        bool IsNeighbor(TVertex source, TVertex dest);

        ISet<TVertex> Vertices { get; }

        IEnumerable<TEdge> Edges { get; }

        bool IsEmpty();

        void RemoveZeroDegreeNodes();

        IEnumerable<TEdge> GetEdges(TVertex source, TVertex dest);

        int GetInDegree(TVertex vertex);

        int GetOutDegree(TVertex vertex);

        IEnumerable<ISet<TVertex>> ConnectedComponents { get; }

    }
}
