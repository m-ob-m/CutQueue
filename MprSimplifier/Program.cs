using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace MprSimplifier
{
    static class Program
    {
        private static Uri InputFileUri { get; set; } = null;
        private static Uri OutputFileUri { get; set; } = null;
        static void Main(string[] arguments)
        {
            ReadArguments(arguments.ToList());
            Mpr.Simplifier.Simplify(new Mpr.File(InputFileUri), new Mpr.File(OutputFileUri));
        }

        private static void ReadArguments(List<string> arguments)
        {
            int i = 0;
            while (i < arguments.Count)
            {
                if (arguments[i] == "--input-file")
                {
                    arguments.RemoveAt(i);
                    if (i < arguments.Count)
                    {
                        if (!new Regex(@"\A--.*\z").IsMatch(arguments[i]))
                        {
                            SetInputFileUri(arguments[i]);
                            arguments.RemoveAt(i);
                        }
                        else
                        {
                            throw new Exception("\"--input-file\" argument is used, but input file path is missing or invalid.");
                        }
                    }
                    else
                    {
                        throw new Exception("\"--input-file\" argument is used, but input file path is missing or invalid.");
                    }
                }
                else if (arguments[i] == "--output-file")
                {
                    arguments.RemoveAt(i);
                    if (i < arguments.Count)
                    {
                        if (!new Regex(@"\A--.*\z").IsMatch(arguments[i]))
                        {
                            SetOutputFileUri(arguments[i]);
                            arguments.RemoveAt(i);
                        }
                        else
                        {
                            throw new Exception("\"--output-file\" argument is used, but output file path is missing or invalid.");
                        }
                    }
                    else
                    {
                        throw new Exception("\"--output-file\" argument is used, but output file path is missing or invalid.");
                    }
                }
                else
                {
                    i++;
                }
            }

            if (InputFileUri == null)
            {
                if (arguments.Count > 0)
                {
                    SetInputFileUri(arguments[0]);
                    arguments.RemoveAt(0);
                }
                else
                {
                    throw new Exception("Input file is not specified.");
                }
            }

            if (OutputFileUri == null)
            {
                if (arguments.Count > 0)
                {
                    SetOutputFileUri(arguments[0]);
                    arguments.RemoveAt(0);
                }
                else
                {
                    throw new Exception("Output file is not specified.");
                }
            }

            if (arguments.Count > 0)
            {
                throw new Exception("Extra arguments were found in the command line.");
            }
        }

        private static void SetInputFileUri(string inputFilePath)
        {
            try
            {
                InputFileUri = new Uri(inputFilePath);
            }
            catch (UriFormatException)
            {
                InputFileUri = new Uri(new Uri(Assembly.GetExecutingAssembly().Location), inputFilePath);
            }
        }

        private static void SetOutputFileUri(string outputFilePath)
        {
            try
            {
                OutputFileUri = new Uri(outputFilePath);
            }
            catch (UriFormatException)
            {
                OutputFileUri = new Uri(new Uri(Assembly.GetExecutingAssembly().Location), outputFilePath);
            }
        }
    }
}
