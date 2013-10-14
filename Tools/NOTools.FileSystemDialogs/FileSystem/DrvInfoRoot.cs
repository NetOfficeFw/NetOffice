using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    internal class DrvInfoRoot : FileSystemInfo
    {
        internal DrvInfoRoot(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false) : base(parent, name, path, loadDirectories, loadFiles)
        {
            Drives = new DrvInfo[0];
        }

        public override bool Exists
        {
            get { return true; ; }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }

        public override bool IsDrivesLoaded { get; internal set; }

        public override bool HasDrives
        {
            get
            {
                if (!IsDrivesLoaded)
                    LoadDrives();
                return Drives.Length > 0;
            }
        }

        public override void LoadFiles()
        {

        }

        public override void LoadDirectories()
        {

        }

        public override void LoadDrives()
        {
            List<DrvInfo> list = new List<DrvInfo>();
            foreach (DriveInfo item in DriveInfo.GetDrives())
            {
                DrvInfo folderInfo = new DrvInfo(Parent, item.Name, CalculateVolumeLabel(item), item.RootDirectory.Root.Name, item.DriveType, item.IsReady, true);
                list.Add(folderInfo);
            }
            Drives = list.ToArray();
            IsDrivesLoaded = true;
        }

        private static string CalculateVolumeLabel(DriveInfo driveInfo)
        {
            string result = "";
            if (driveInfo.DriveType == DriveType.Network)
            {
                if (!driveInfo.IsReady)
                    return "";

                string rootDirectory = driveInfo.RootDirectory.Root.Name;
                result = ResolveShareName(driveInfo.RootDirectory.Root.Name);

                if (rootDirectory.EndsWith("\\", StringComparison.InvariantCultureIgnoreCase))
                    rootDirectory = rootDirectory.Substring(0, rootDirectory.Length - 1);

                if (result.StartsWith("(" + rootDirectory + ")", StringComparison.InvariantCultureIgnoreCase))
                    result = result.Substring(("(" + rootDirectory + ")").Length);
            }
            else if (driveInfo.DriveType == DriveType.CDRom)
            {
                if (driveInfo.IsReady)
                    result = driveInfo.VolumeLabel;
                else
                    return " (CD-ROM)";
            }
            else
            {
                if (!driveInfo.IsReady)
                    return "";
                result =  driveInfo.VolumeLabel;
            }

            result = result.Trim();
            if (!String.IsNullOrWhiteSpace(result))
                result = " (" + result + ")";
            return result;
        }

        private static string ResolveShareName(string path)
        {
            Win32.SHFILEINFO shfi = new Win32.SHFILEINFO();
            Win32.SHGFI dwflag = Win32.SHGFI.SHGFI_DISPLAYNAME | Win32.SHGFI.SHGFI_TYPENAME;
            int dwAttr = 0;
            dwflag = dwflag | Win32.SHGFI.SHGFI_USEFILEATTRIBUTES;
            dwAttr = 0x80;
            Win32.SHGetFileInfo(path, dwAttr, ref shfi, Win32.cbFileInfo, Convert.ToInt32(dwflag));
            string result = shfi.szDisplayName;

            return result;
        }
    }
}
