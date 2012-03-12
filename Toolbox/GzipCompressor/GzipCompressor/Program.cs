using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace GzipCompressor
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Compress xml files to gzip.");          

            DirectoryInfo di = new DirectoryInfo(Environment.CurrentDirectory);
            foreach (FileInfo fi in di.GetFiles("*.dll"))
            {
                using (FileStream inFile = fi.OpenRead())
                {
                    if ((File.GetAttributes(fi.FullName) & FileAttributes.Hidden) != FileAttributes.Hidden & fi.Extension != ".gz")
                    {
                        // Create the compressed file.
                        using (FileStream outFile = File.Create(fi.FullName + ".gz"))
                        {
                            using (GZipStream Compress = new GZipStream(outFile, CompressionMode.Compress))
                            {
                                inFile.CopyTo(Compress);
                                Console.WriteLine("Compressed {0} from {1} to {2} bytes.", fi.Name, fi.Length.ToString(), outFile.Length.ToString());
                            }
                        }
                    }
                }
            }
        }
    }
}
