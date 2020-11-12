using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace MprSimplifier.Mpr
{
    class Variable
    {
        public string Name { get; internal set; }
        public string Value { get; private set; }
        public string Description { get; internal set; }
        public HashSet<string> DependsOn { get; internal set; }

        private Variable(string name, string value, string description, HashSet<string> dependsOn)
        {
            Name = name;
            Value = value;
            Description = description;
            DependsOn = dependsOn;
        }

        public static Variable FromText(string text)
        {
            Match match = new Regex(
                @"\A(?<NAME>.*?)=""(?<VALUE>.*)""\r\nKM=""(?<DECRIPTION>.*)""(?:\r\n)?\z"
            ).Match(text);

            if (!match.Success)
            {
                throw new Exception("Invalid mpr variable text provided.");
            }

            MatchCollection dependsOnMatches = new Regex(@"\b[A-Za-z_][A-Za-z0-9_]*\b")
                .Matches(match.Groups["VALUE"].Value);
            HashSet<string> dependsOn = new HashSet<string>();
            foreach (Match dependsOnMatch in dependsOnMatches)
            {
                dependsOn.Add(dependsOnMatch.Value);
            }

            return new Variable(
                match.Groups["NAME"].Value,
                match.Groups["VALUE"].Value,
                match.Groups["DESCRIPTION"].Value,
                dependsOn
            );
        }

        public override string ToString()
        {
            return $"{Name}=\"{Value}\"\r\nKM=\"{Description}\"\r\n";
        }
    }
}
