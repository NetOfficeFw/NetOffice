using System;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.IO.Compression;
using System.Text;
using NOTools.CSharpTextEditor;

namespace NOTools.CSharpTextEditorPFCreator
{
    internal static class PersistenceFileCreator
    {
        internal static void CreatePersistenceFiles(string resultFolder, string[] assemblies, bool createCompressedCopies)
        {
            if (!Directory.Exists(resultFolder))
                Directory.CreateDirectory(resultFolder);

            if(Directory.EnumerateFiles(resultFolder).Count() > 0)
                Directory.Delete(resultFolder, true);

            CodeEditorControl editor = new CodeEditorControl();
            editor.PersistencePath = resultFolder;

            foreach (string item in assemblies)
            {
                string name = System.IO.Path.GetFileNameWithoutExtension(item);
                editor.AddReferenceFromFile(name, item);
            }

            if (createCompressedCopies)
                CompressFolderFiles(resultFolder);

            editor.Dispose();
        }

        internal static void CompressFolderFiles(string resultFolder)
        {
            DirectoryInfo directorySelected = new DirectoryInfo(resultFolder);
            foreach (FileInfo fileToCompress in directorySelected.GetFiles("*.*", SearchOption.AllDirectories))
                CompressFile(fileToCompress);
           
        }

        private static void CompressFile(FileInfo fileToCompress)
        {
            using (FileStream originalFileStream = fileToCompress.OpenRead())
            {
                if ((File.GetAttributes(fileToCompress.FullName) & FileAttributes.Hidden) != FileAttributes.Hidden & fileToCompress.Extension != ".gz")
                {
                    using (FileStream compressedFileStream = File.Create(fileToCompress.FullName + ".gz"))
                    {
                        using (GZipStream compressionStream = new GZipStream(compressedFileStream, CompressionMode.Compress))
                        {
                            originalFileStream.CopyTo(compressionStream);
                        }
                    }
                }
            }
        }
    }
}
