using System;
using System.IO;
using System.Threading;

namespace CutQueue.Lib.Tools
{
    static class FileAdvanced
    {
        private const uint MAXIMUM_WAIT_TIME_FOR_DELETION_IN_SECONDS = 5;

        /// <summary>
        /// Deletes a file.
        /// </summary>
        /// <param name="path">The path of the file.</param>
        public static void Delete(string path)
        {
            using (ManualResetEventSlim manualResetEventSlim = new ManualResetEventSlim())
            {
                if (File.Exists(path))
                {
                    using (FileSystemWatcher fileSystemWatcher = new FileSystemWatcher(Path.GetDirectoryName(path)))
                    {
                        manualResetEventSlim.Reset();

                        fileSystemWatcher.Filter = Path.GetFileName(path);
                        fileSystemWatcher.Deleted += (object source, FileSystemEventArgs e) => { manualResetEventSlim.Set(); };
                        fileSystemWatcher.NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite;
                        fileSystemWatcher.EnableRaisingEvents = true;
                        File.Delete(path);

                        manualResetEventSlim.Wait((int)MAXIMUM_WAIT_TIME_FOR_DELETION_IN_SECONDS * 1000);
                        if (!manualResetEventSlim.IsSet)
                        {
                            throw new Exception($"File \"{path}\" took more than the maximum allowed time of {MAXIMUM_WAIT_TIME_FOR_DELETION_IN_SECONDS} seconds to delete.");
                        }
                    }
                }
                else if (Directory.Exists(path))
                {
                    using (FileSystemWatcher fileSystemWatcher = new FileSystemWatcher(Path.GetDirectoryName(path)))
                    {
                        manualResetEventSlim.Reset();

                        fileSystemWatcher.Filter = Path.GetFileName(path);
                        fileSystemWatcher.Deleted += (object source, FileSystemEventArgs e) => { manualResetEventSlim.Set(); };
                        fileSystemWatcher.NotifyFilter = NotifyFilters.DirectoryName | NotifyFilters.LastWrite;
                        fileSystemWatcher.EnableRaisingEvents = true;
                        Directory.Delete(path);

                        manualResetEventSlim.Wait((int)MAXIMUM_WAIT_TIME_FOR_DELETION_IN_SECONDS * 1000);
                        if (!manualResetEventSlim.IsSet)
                        {
                            throw new Exception($"Folder \"{path}\" took more than the maximum allowed time of {MAXIMUM_WAIT_TIME_FOR_DELETION_IN_SECONDS} seconds to delete.");
                        }
                    }
                }
            }
        }
    }
}
