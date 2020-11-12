using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MprSimplifier.Mpr
{
    class CoordinateSystem
    {
        private string index;
        public string Index
        {
            get { return index; }
            internal set
            {
                index = value;
                Data = new Regex(@"(?<=^NR="")[A-Z0-9]+(?=""\r\n)", RegexOptions.Multiline).Replace(Data, index, 1);
            }
        }
        public HashSet<string> DependsOn { get; private set; }
        public string Data { get; private set; }

        private CoordinateSystem(string data, string index, HashSet<string> dependsOn)
        {
            Data = data.Trim();
            this.index = index;
            DependsOn = dependsOn;
        }

        public static CoordinateSystem FromText(string data)
        {
            Match indexMatch = new Regex(@"(?<=^NR="")[A-Z0-9]+(?=""\r\n)", RegexOptions.Multiline).Match(data);
            if (!indexMatch.Success)
            {
                throw new Exception("Cannot find coordinate system index");
            }
            string index = indexMatch.Value;

            MatchCollection dependsOnMatches = new Regex(
                @"(?<=^[A-Z]+=.*)\b[A-Za-z_][A-Za-z0-9_]*\b",
                RegexOptions.Multiline
            ).Matches(data);
            HashSet<string> dependsOn = new HashSet<string>() { "_BSX", "_BSY" };
            foreach (Match dependsOnMatch in dependsOnMatches)
            {
                dependsOn.Add(dependsOnMatch.Value);
            }

            return new CoordinateSystem(data, index, dependsOn);
        }

        public override string ToString()
        {
            return Data;
        }
    }
}
