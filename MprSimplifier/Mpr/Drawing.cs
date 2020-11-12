using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MprSimplifier.Mpr
{
    class Drawing
    {
        private uint index;
        private string coordinateSystem;
        public uint Index
        {
            get { return index; }
            internal set
            {
                index = value;
                Data = new Regex(@"(?<=\A\])\d+(?=\r\n)").Replace(Data, index.ToString(), 1);
            }
        }

        public string CoordinateSystem
        {
            get { return coordinateSystem; }
            internal set
            {
                coordinateSystem = value;
                Data = new Regex(@"(?<=^KO=)[A-Z0-9]+(?=\r\n)", RegexOptions.Multiline).Replace(Data, index.ToString(), 1);
            }
        }
        public HashSet<string> DependsOn { get; private set; }
        public string Data { get; private set; }

        private Drawing(string data, uint index, string coordinateSystem, HashSet<string> dependsOn)
        {
            Data = data.Trim();
            this.index = index;
            this.coordinateSystem = coordinateSystem;
            DependsOn = dependsOn;
        }

        public static Drawing FromText(string data)
        {
            Match indexMatch = new Regex(@"(?<=\A\])\d+(?=\r\n)").Match(data);
            if (!indexMatch.Success)
            {
                throw new Exception("Cannot find drawing index.");
            }
            uint index = uint.Parse(indexMatch.Value);

            Match coordinateSystemMatch = new Regex(@"(?<=^KO=)[A-Z0-9]+(?=\r\n)", RegexOptions.Multiline)
                .Match(data);
            if (!coordinateSystemMatch.Success)
            {
                throw new Exception("Cannot find coordinate system for drawing.");
            }
            string coordinateSystem = coordinateSystemMatch.Value;

            MatchCollection dependsOnMatches = new Regex(
                @"(?<=^[A-Z]+=.*)\b[A-Za-z_][A-Za-z0-9_]*\b",
                RegexOptions.Multiline
            ).Matches(data);
            HashSet<string> dependsOn = new HashSet<string>();
            foreach (Match dependsOnMatch in dependsOnMatches)
            {
                dependsOn.Add(dependsOnMatch.Value);
            }

            return new Drawing(data, index, coordinateSystem, dependsOn);
        }

        public override string ToString()
        {
            return Data;
        }
    }
}
