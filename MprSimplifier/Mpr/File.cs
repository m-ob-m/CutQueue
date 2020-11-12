using System;
using System.IO;
using System.Text;

namespace MprSimplifier.Mpr
{
    class File
    {
        public Uri Uri { get; internal set; }

        public string Data { get; internal set; }

        public File(Uri uri)
        {
            Uri = uri;
        }

        public File Read()
        {
            if (!Exists())
            {
                throw new FileNotFoundException($"Mpr file \"{Uri.LocalPath}\" does not exist.");
            }

            Data = System.IO.File.ReadAllText(Uri.LocalPath, Encoding.GetEncoding("iso-8859-1"));
            return this;
        }

        private bool Exists()
        {
            return System.IO.File.Exists(Uri.LocalPath);
        }

        public File Write()
        {
            System.IO.File.WriteAllText(Uri.LocalPath, Data, Encoding.GetEncoding("iso-8859-1"));
            return this;
        }
    }
}
