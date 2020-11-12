using CutQueue.Lib.Exceptions;
using System.Diagnostics;

namespace CutQueue.Lib.Tools
{
    static class ProcessAdvanced
    {
        /// <summary>
        /// Executes a process.
        /// </summary>
        /// <param name="processName">The filepath of the process to execute</param>
        /// <param name="arguments">A string of arguments used to start the process</param>
        /// <param name="maximumExecutionTime">The maximum execution time of the process in seconds</param>
        /// <returns>The process' output</returns>
        public static (int exitCode, string standardOutput, string standardError) ExecuteProcess(string processName, string arguments, int maximumExecutionTime = 60)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo
            {
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                FileName = processName,
                Arguments = arguments,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                UseShellExecute = false
            };
            using (Process process = Process.Start(processStartInfo))
            {
                string standardOutput = process.StandardOutput.ReadToEnd();
                string standardError = process.StandardError.ReadToEnd();
                process.WaitForExit(maximumExecutionTime * 1000);
                if (!process.HasExited)
                {
                    process.Kill();
                    throw new MaximumProcessExecutionTimeReachedException(processName, arguments, maximumExecutionTime);
                }

                return (process.ExitCode, standardOutput, standardError);
            }
        }
    }
}
