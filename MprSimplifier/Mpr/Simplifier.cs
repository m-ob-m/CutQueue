using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace MprSimplifier.Mpr
{
    class Simplifier
    {
        public File InputFile { get; internal set; }
        public File OutputFile { get; internal set; }
        private Header Header { get; set; }
        private List<Variable> Variables { get; set; }
        private List<CoordinateSystem> CoordinateSystems { get; set; }
        private List<Drawing> Drawings { get; set; }
        private List<Machining> Machinings { get; set; }
        private string Buffer { get; set; }

        private Simplifier(File inputFile, File outputFile)
        {
            InputFile = inputFile;
            OutputFile = outputFile;
        }

        private Simplifier Parse()
        {
            Debug.Write("Start input file parsing...\r\n");

            Buffer = InputFile.Read().Data;
            ExtractHeader();
            ExtractClosure();
            ExtractMachinings();
            ExtractVariables();
            ExtractCoordinateSystems();
            ExtractDrawings();
            if (!string.IsNullOrEmpty(Buffer))
            {
                throw new Exception("Data remains in buffer after parsing.");
            }

            Debug.Write("End input file parsing...\r\n");

            return this;
        }

        private Simplifier ExtractHeader()
        {
            Debug.Write("Start header extraction...\r\n");

            Match match = new Regex(@"\A\[H\r\n(?:.+\r\n)+\r\n").Match(Buffer);
            Header = new Header(match.Value);
            Buffer = Buffer.Substring(match.Length);

            Debug.Write("End header extraction...\r\n");

            return this;
        }

        private Simplifier ExtractVariables()
        {
            Debug.Write("Start variables extraction...\r\n");

            Match variablesMatch = new Regex(@"\A\[001\r\n(?<VARIABLES>(?:.+\r\n)*)\r\n").Match(Buffer);
            Buffer = Buffer.Substring(variablesMatch.Length);

            Variables = new List<Variable>();
            if (!string.IsNullOrEmpty(variablesMatch.Groups["VARIABLES"].Value))
            {
                MatchCollection matches = new Regex(@"(?<=\A|\G).*\r\n.*\r\n")
                    .Matches(variablesMatch.Groups["VARIABLES"].Value);

                foreach (Match match in matches)
                {
                    Variables.Add(Variable.FromText(match.Value));
                }
            }

            Debug.Write("End variables extraction...\r\n");

            return this;
        }

        private Simplifier ExtractCoordinateSystems()
        {
            Debug.Write("Start planes extraction...\r\n");

            Match coordinateSystemsMatch = new Regex(@"\A\[K\r\n(?<COORDINATE_SYSTEMS>(?:<00 .*\r\n(?:.+\r\n)+\r\n)*)")
                .Match(Buffer);
            Buffer = Buffer.Substring(coordinateSystemsMatch.Length);

            CoordinateSystems = new List<CoordinateSystem>();
            if (!string.IsNullOrEmpty(coordinateSystemsMatch.Groups["COORDINATE_SYSTEMS"].Value))
            {
                MatchCollection matches = new Regex(@"(?<=\A|\G)<00 .*\r\n(?:.+\r\n)+\r\n")
                    .Matches(coordinateSystemsMatch.Groups["COORDINATE_SYSTEMS"].Value);

                foreach (Match match in matches)
                {
                    CoordinateSystems.Add(CoordinateSystem.FromText(match.Value));
                }
            }

            Debug.Write("End planes extraction...\r\n");

            return this;
        }

        private Simplifier ExtractDrawings()
        {
            Debug.Write("Start drawings extraction...\r\n");

            Regex regularExpression = new Regex(@"\A\](?<INDEX>\d+)\r\n(?:.*(\r\n|\z))+?(?=]\d+|\z)");

            Match match;
            Drawings = new List<Drawing>();
            while ((match = regularExpression.Match(Buffer)).Success)
            {
                Drawings.Add(Drawing.FromText(match.Value));
                Buffer = Buffer.Substring(match.Length);
            }

            Debug.Write("End drawings extraction...\r\n");

            return this;
        }

        private Simplifier ExtractMachinings()
        {
            Debug.Write("Start machinings extraction...\r\n");

            Regex regularExpression = new Regex(@"<(?!00 )\d+ .*?\r\n(?:.+\r\n)+.+\z", RegexOptions.RightToLeft);

            Match match;
            Machinings = new List<Machining>();
            while ((match = regularExpression.Match(Buffer)).Success)
            {
                Machinings.Add(Machining.FromText(match.Value));
                Buffer = Buffer.Substring(0, match.Index - 1);
            }
            Machinings.Reverse();

            Debug.Write("End machinings extraction...\r\n");

            return this;
        }

        private Simplifier ExtractClosure()
        {
            Debug.Write("Start closure extraction...\r\n");

            Match match = new Regex(@"(?:\r\n)*!(?:\r\n)*\z", RegexOptions.RightToLeft).Match(Buffer);
            if (!match.Success)
            {
                throw new Exception("Closure is missing from input mpr file.");
            }
            Buffer = Buffer.Substring(0, match.Index - 1);

            Debug.Write("End closure extraction...\r\n");

            return this;
        }

        public static void Simplify(File inputFile, File outputFile)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            Simplifier simplifier = new Simplifier(inputFile, outputFile);
            simplifier.Parse();
            simplifier.RemoveUnusedElements();
            simplifier.OutputFile.Data = simplifier.ToString();
            simplifier.OutputFile.Write();
            stopwatch.Stop();
            double duration = Math.Round((double)stopwatch.ElapsedMilliseconds / 1000, 0, MidpointRounding.AwayFromZero);
            Debug.Write($"Simplification process lasted {duration} seconds.\r\n");
        }

        private Simplifier RemoveUnusedElements()
        {
            Debug.Write("Start unused elements removal...\r\n");

            RemoveUnusedMachinings();
            RemoveUnusedDrawings();
            RemoveUnusedCoordinateSystems();
            RemoveUnusedVariables();
            Verify();

            Debug.Write("End unused elements removal...\r\n");

            return this;
        }

        private Simplifier RemoveUnusedMachinings()
        {
            Debug.Write("Start unused machinings removal...\r\n");

            for (int i = Machinings.Count - 1; i >= 0; i--)
            {
                if (!new List<uint>() { 100, 101, 105, 112, 181 }.Contains(Machinings[i].Code))
                {
                    Debug.Write($"Non supported machining code {Machinings[i].Code} found.\r\n");
                    throw new Exception($"Non supported machining code {Machinings[i].Code} found.");
                }
                if (Machinings[i].Enable == false)
                {
                    Machinings.RemoveAt(i);
                }
            }

            Debug.Write("End unused machinings removal...\r\n");

            return this;
        }

        private Simplifier RemoveUnusedDrawings()
        {
            Debug.Write("Start unused drawings removal...\r\n");

            for (int i = Drawings.Count - 1; i >= 0; i--)
            {
                Drawing drawing = Drawings[i];
                bool isUsed = false;

                foreach (Machining machining in Machinings)
                {
                    if (machining.Drawing == drawing.Index)
                    {
                        isUsed = true;
                        break;
                    }
                }

                if (!isUsed)
                {
                    Drawings.RemoveAt(i);

                    foreach (Machining machining in Machinings)
                    {
                        if (machining.Drawing > drawing.Index)
                        {
                            machining.Drawing--;
                        }
                        else if (machining.Drawing == drawing.Index)
                        {
                            throw new Exception("Removed a drawing that still had associated machinings.");
                        }
                    }

                    for (int j = Drawings.Count - 1; j >= i; j--)
                    {
                        Drawings[j].Index--;
                    }
                }
            }

            Debug.Write("End unused drawings removal...\r\n");

            return this;
        }

        private Simplifier RemoveUnusedCoordinateSystems()
        {
            Debug.Write("Start unused coordinate systems removal...\r\n");

            for (int i = CoordinateSystems.Count - 1; i >= 0; i--)
            {
                bool isUsed = false;
                foreach (Machining machinining in Machinings)
                {
                    if (machinining.CoordinateSystem == CoordinateSystems[i].Index)
                    {
                        isUsed = true;
                        break;
                    }
                }

                if (!isUsed)
                {
                    foreach (Drawing drawing in Drawings)
                    {
                        if (drawing.CoordinateSystem == CoordinateSystems[i].Index)
                        {
                            isUsed = true;
                            break;
                        }
                    }
                }

                if (!isUsed)
                {
                    CoordinateSystems.RemoveAt(i);
                }
            }

            Debug.Write("End unused coordinate systems removal...\r\n");

            return this;
        }

        private Simplifier RemoveUnusedVariables()
        {
            Debug.Write("Start unused variables removal...\r\n");

            HashSet<string> usedVariables = new HashSet<string>();
            foreach (Machining machining in Machinings)
            {
                usedVariables.UnionWith(machining.DependsOn);
            }
            foreach (Drawing drawing in Drawings)
            {
                usedVariables.UnionWith(drawing.DependsOn);
            }
            foreach (CoordinateSystem coordinateSystem in CoordinateSystems)
            {
                usedVariables.UnionWith(coordinateSystem.DependsOn);
            }

            uint amountTreated = 0;
            uint totalAmount = (uint)Variables.Count;
            int i = Variables.Count - 1;
            while (i >= 0)
            {
                if (!usedVariables.Contains(Variables[i].Name))
                {
                    Variables.RemoveAt(i);
                }
                else 
                {
                    usedVariables.UnionWith(Variables[i].DependsOn);
                }

                if (++amountTreated % 500 == 0)
                {
                    Debug.Write($"{Math.Round((double)amountTreated / totalAmount * 100, 0, MidpointRounding.AwayFromZero)}%\r\n");
                }
                
                i--;
            }

            Debug.Write("End unused variables removal...\r\n");

            return this;
        }

        private Simplifier Verify()
        {
            Debug.Write("Start verification...\r\n");

            foreach (Machining machining in Machinings)
            {
                bool? coordinateSystemFound = machining.CoordinateSystem == null ? (bool?)null : false;
                
                if (coordinateSystemFound == false && new List<string>() { "00", "A00", "B00", "C00", "D00" }.Contains(machining.CoordinateSystem))
                {
                    coordinateSystemFound = true;
                }

                if (coordinateSystemFound == false)
                {
                    foreach (CoordinateSystem coordinateSystem in CoordinateSystems)
                    {
                        if (coordinateSystem.Index == machining.CoordinateSystem)
                        {
                            coordinateSystemFound = true;
                            break;
                        }
                    }
                }

                if (coordinateSystemFound == false)
                {
                    throw new Exception($"Post simplification verification failed: coordinate system {machining.CoordinateSystem} is missing.");
                }

                bool? drawingFound = machining.Drawing == null ? (bool?)null : false;

                if (drawingFound == false)
                {
                    foreach (Drawing drawing in Drawings)
                    {
                        if (drawing.Index == machining.Drawing)
                        {
                            drawingFound = true;
                            break;
                        }
                    }

                    if (drawingFound == false)
                    {
                        throw new Exception("Post simplification verification failed: A drawing is missing.");
                    }
                }
            }

            foreach (Drawing drawing in Drawings)
            {
                bool? coordinateSystemFound = drawing.CoordinateSystem == null ? (bool?)null : false;

                if (coordinateSystemFound == false && new List<string>() { "00", "A00", "B00", "C00", "D00" }.Contains(drawing.CoordinateSystem))
                {
                    coordinateSystemFound = true;
                }

                if (coordinateSystemFound == false)
                {
                    foreach (CoordinateSystem coordinateSystem in CoordinateSystems)
                    {
                        if (coordinateSystem.Index == drawing.CoordinateSystem)
                        {
                            coordinateSystemFound = true;
                            break;
                        }
                    }
                }

                if (coordinateSystemFound == false)
                {
                    throw new Exception($"Post simplification verification failed: coordinate system {drawing.CoordinateSystem} is missing.");
                }
            }

            Debug.Write("End verification...\r\n");

            return this;
        }

        public override string ToString()
        {
            string mpr = Header.Data + "\r\n\r\n";

            if (Variables.Count > 0)
            {
                mpr += "[001\r\n";
            }

            foreach (Variable variable in Variables)
            {
                mpr += variable.ToString();
            }

            mpr += "\r\n";

            if (CoordinateSystems.Count > 0)
            {
                mpr += "[K\r\n";
            }

            foreach (CoordinateSystem coordinateSystem in CoordinateSystems)
            {
                mpr += coordinateSystem.ToString() + "\r\n\r\n";
            }

            foreach (Drawing drawing in Drawings)
            {
                mpr += drawing.ToString() + "\r\n\r\n";
            }

            for (int i = 0; i < Machinings.Count; i++)
            {
                mpr += Machinings[i].ToString() + "\r\n";

                if (i != Machinings.Count - 1)
                {
                    mpr += "\r\n";
                }
            }

            mpr += "!\r\n";

            return mpr;
        }
    }
}
