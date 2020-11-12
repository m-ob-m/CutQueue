using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MprSimplifier.Mpr
{
    class Machining
    {
        private uint code;
        private uint? drawing;
        private string coordinateSystem;
        private bool? enable;
        public bool? Enable
        {
            get { return enable; }
            set
            {
                enable = value;
                if (enable == null)
                {
                    new Regex(@"\r\nEN=""[01]""").Replace(Data, "");
                }
                else
                {
                    Regex regularExpression = new Regex(@"(?<=^EN="")\d+(?=""\r\n|\z)", RegexOptions.Multiline);
                    if (regularExpression.IsMatch(Data))
                    {
                        Data = regularExpression.Replace(Data, (bool)enable ? "1" : "0");
                    }
                    else
                    {
                        Data += "\r\nEN=\"" + ((bool)enable ? "1" : "0") + "\"";
                    }
                }
            }
        }
        public uint Code
        {
            get { return code; }
            private set
            {
                code = value;
                Data = new Regex(@"(?<=\A<)\d+").Replace(Data, code.ToString());
            }
        }
        public uint? Drawing
        {
            get { return drawing; }
            internal set
            {
                drawing = value;
                Data = new Regex(@"(?<=^E[AE]="")\d+(?=:\d+""\r\n)", RegexOptions.Multiline).Replace(Data, drawing.ToString());
            }
        }
        public string CoordinateSystem
        {
            get { return coordinateSystem; }
            set
            {
                coordinateSystem = value;
                Data = new Regex(@"(?<=^KO="")[A-Z0-9]+(?=""\r\n)", RegexOptions.Multiline).Replace(Data, drawing.ToString());
            }
        }
        public HashSet<string> DependsOn { get; private set; }
        public string Data { get; private set; }

        private Machining(string data, uint code, string coordinateSystem, uint? drawing, bool? enable, HashSet<string> dependsOn)
        {
            Data = data.Trim();
            this.code = code;
            this.coordinateSystem = coordinateSystem;
            this.drawing = drawing;
            this.enable = enable;
            DependsOn = dependsOn;
        }

        public static Machining FromText(string data)
        {
            Match codeMatch = new Regex(@"(?<=\A<)\d+").Match(data);
            if (!codeMatch.Success)
            {
                throw new Exception("Cannot find machining code.");
            }
            uint code = uint.Parse(codeMatch.Value);

            Match drawingMatch = new Regex(@"(?<=^E[AE]="")\d+(?=:\d+""\r\n)", RegexOptions.Multiline).Match(data);
            uint? drawing = null;
            if (drawingMatch.Success)
            {
                drawing = uint.Parse(drawingMatch.Value);
            }

            Match coordinateSystemMatch = new Regex(@"(?<=^KO="")[A-Z0-9]+(?=""\r\n)", RegexOptions.Multiline).Match(data);
            string coordinateSystem = null;
            if (coordinateSystemMatch.Success)
            {
                coordinateSystem = coordinateSystemMatch.Value;
            }

            Match enableMatch = new Regex(@"(?<=^EN="")\d+(?=""\r\n|\z)", RegexOptions.Multiline).Match(data);
            bool? enable;
            if (enableMatch.Success)
            {
                if (enableMatch.Value == "1")
                {
                    enable = true;
                }
                else if (enableMatch.Value == "0")
                {
                    enable = false;
                }
                else
                {
                    throw new Exception("Invalid value for machining enable bit.");
                }
            }
            else
            {
                enable = null;
            }

            MatchCollection dependsOnMatches = new Regex(
                @"(?<=^(?!MNM|KAT|KM)[^=]+=.*)\b[A-Za-z_][A-Za-z0-9_]*\b",
                RegexOptions.Multiline
            ).Matches(data);
            HashSet<string> dependsOn = new HashSet<string>();
            foreach (Match dependsOnMatch in dependsOnMatches)
            {
                dependsOn.Add(dependsOnMatch.Value);
            }

            return new Machining(data, code, coordinateSystem, drawing, enable, dependsOn);
        }

        public override string ToString()
        {
            return Data;
        }
    }
}
