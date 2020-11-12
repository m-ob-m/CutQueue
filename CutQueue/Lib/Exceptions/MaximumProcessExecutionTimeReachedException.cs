using System;
using System.IO;

namespace CutQueue.Lib.Exceptions
{
    class MaximumProcessExecutionTimeReachedException : Exception
    {
        public MaximumProcessExecutionTimeReachedException(string processPath, string argumentString, int maximumExecutionTime) :
                base($"Process {Path.GetFileName(processPath)} took longer than {maximumExecutionTime} seconds to execute.")
        {
            Data.Add("ProcessPath", processPath);
            Data.Add("ArgumentString", argumentString);
            Data.Add("MaximumExecutionTime", maximumExecutionTime);
        }
    }
}
