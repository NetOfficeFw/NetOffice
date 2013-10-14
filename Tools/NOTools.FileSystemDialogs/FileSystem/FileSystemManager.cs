using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    internal class FileSystemManager
    {
        internal string FileSearchPattern { get; set; }

        internal FolderInfo GetDesktop(bool load = false)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string name = Path.GetFileName(folderPath);
            FolderInfo folderInfo = new FolderInfo(this, name, folderPath, load);
            return folderInfo;
        }

        internal DrvInfoRoot GetMyComputer()
        {
            DrvInfoRoot driveRoot = new DrvInfoRoot(this, "MyComputer","");
            return driveRoot;
        }

        internal FolderInfo GetMyDocuments(bool load = false)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string name = Path.GetFileName(folderPath);
            FolderInfo folderInfo = new FolderInfo(this, name, folderPath, load);
            return folderInfo;
        }

        internal SpecialFolderRoot GetSpecialFolders()
        {
            SpecialFolderRoot driveRoot = new SpecialFolderRoot(this, "SpecialFolders", "");
            return driveRoot;
        }

        internal TemplateFolderRoot GetTemplateFolders(TemplateFolderDescription[] description)
        {
            TemplateFolderRoot folderRooth = new TemplateFolderRoot(this, "TemplateFolders", "", description);
            return folderRooth;
        }
    }
}
