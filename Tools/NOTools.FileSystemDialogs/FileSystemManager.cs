using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public abstract class FileSystemInfo
    {
        internal FileSystemInfo(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false)
        {
            Parent = parent;
            Name = name;
            Path = path;
            Drives = new DrvInfo[0];
            Directories = new FolderInfo[0];
            Files = new FiInfo[0];
            if (loadDirectories)
                LoadDirectories();
            if (loadFiles)
                LoadFiles();
        }

        /// <summary>
        /// Parent FileSystemManager instance
        /// </summary>
        internal FileSystemManager Parent { get; private set; }

        /// <summary>
        /// Display name 
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Full qualified path
        /// </summary>
        public string Path { get; private set; }

        /// <summary>
        /// Returns the instance exists currently in the file system
        /// </summary>
        public abstract bool Exists { get; }

        /// <summary>
        /// Returns the instance is already loaded from the filesystem
        /// </summary>
        public abstract bool IsDirectoriesLoaded { get; internal set; }

        /// <summary>
        /// Returns the instance is already loaded from the filesystem
        /// </summary>
        public abstract bool IsFilesLoaded { get; internal set; }

        public abstract bool IsDrivesLoaded { get; internal set; }
         
        /// <summary>
        /// Returns the corresponding item in the filesystem is currently not ready to use
        /// </summary>
        public virtual bool IsReady
        {
            get
            {
                return true;
            }
            internal
            set
            { 
            
            }
        }

        /// <summary>
        /// Returns the info an access error has occured for the instance
        /// </summary>
        public virtual bool HasErrors { get; internal set; }

        /// <summary>
        /// Returns the instance has an own display image
        /// </summary>
        public virtual bool SupportsDisplayImage { get; internal set; }
        
        /// <summary>
        /// Display Image(16x16) (if SupportsDisplayImage is okay)
        /// </summary>
        public virtual Icon DisplayImageSmall{ get; internal set; }

        /// <summary>
        /// Display Image(32x32) (if SupportsDisplayImage is okay)
        /// </summary>
        public virtual Icon DisplayImageLarge { get; internal set; }


        public DrvInfo[] Drives
        {
            get
            {
                if (!IsDrivesLoaded)
                    LoadDrives();
                return _drives;
            }
            internal set
            {
                _drives = value;
            }
        }
        private DrvInfo[] _drives;

        /// <summary>
        /// First level sub directory(if instance is a directory)
        /// </summary>
        public virtual FolderInfo[] Directories 
        {
            get
            {
                if (!IsDirectoriesLoaded)
                    LoadDirectories();
                return _directories;
            }
            internal set
            {
                _directories = value;
            }
        }
        private FolderInfo[] _directories;

        /// <summary>
        /// All files (if instance is a directory or drive)
        /// </summary>
        public virtual FiInfo[] Files 
        {
            get
            {
                if (!IsFilesLoaded)
                    LoadFiles();
                return _files;
            }
            internal set
            {
                _files = value;
            }
        }
        private FiInfo[] _files;

        /// <summary>
        /// Returns the instance has first level files
        /// </summary>
        public virtual bool HasFiles
        {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Returns the instance has first level sub directories
        /// </summary>
        public virtual bool HasDirectories
        {
            get
            {
                return false;
            }
        }

        public virtual bool HasDrives
         {
            get
            {
                return false;
            }
        }

        /// <summary>
        /// Load or reload infos from filesystem
        /// </summary>
        public abstract void LoadFiles();

        /// <summary>
        /// Load or reload infos from filesystem
        /// </summary>
        public abstract void LoadDirectories();

        public abstract void LoadDrives();

        //protected virtual bool HasFiles
        //{
        //    get
        //    {
        //        if (!IsLoaded)
        //            RefreshFiles();
        //        return Files.Count() > 0;
        //    }
        //}

        //protected virtual bool HasDirectories
        //{
        //    get
        //    {
        //        if (!IsLoaded)
        //            RefreshDirectories();
        //        return Directories.Count() > 0;
        //    }
        //}

        //{
        //    get 
        //    {
        //        return Directory.Exists(Path);
        //    }
        //}
        //internal string Name
        //{
        //    get
        //    {
        //        string result = System.IO.Path.GetDirectoryName(Path);
        //        if (null != result)
        //            return result;

        //        result = Path;
        //        if (result.EndsWith("\\", StringComparison.InvariantCultureIgnoreCase))
        //            result = result.Substring(0, result.Length - 1);

        //        return result;
        //    }
        //}

         
        //public virtual void RefreshDirectories()
        //{
 
        //}

        //public virtual void RefreshFiles()
        //{ 
        
        //}

        //public virtual void Refresh()
        //{
        //    try
        //    {
        //        if (Exists)
        //        {
        //            DirectoryInfo dirInfo = new DirectoryInfo(Path); 
        //            Directories = dirInfo.GetDirectories("*.*", SearchOption.TopDirectoryOnly);
        //            Files = dirInfo.GetFiles(String.IsNullOrWhiteSpace(Parent.FileSearchPattern) != true ? Parent.FileSearchPattern : "*.*", SearchOption.TopDirectoryOnly);
        //        }
        //        else 
        //        {
        //            Directories = new DirectoryInfo[0];
        //            Files = new FileInfo[0];
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //    finally
        //    {
        //        IsLoaded = true;
        //    }
        //}
    }

    internal class DrvInfoRoot : FileSystemInfo
    { 
        internal DrvInfoRoot (FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false) : base(parent, name, path, loadDirectories, loadFiles)
        {
            Drives = new DrvInfo[0];
        }

        public override bool Exists
        {
            get { return true; ; }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }
     
        public override bool IsFilesLoaded{ get; internal set; }
       
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
                DrvInfo folderInfo = new DrvInfo(Parent, item.Name, item.RootDirectory.Root.Name, item.DriveType, item.IsReady, true);
                list.Add(folderInfo);
            }
            Drives = list.ToArray();
            IsDrivesLoaded = true;
        }
    }

    public class DrvInfo : FileSystemInfo
    {
        internal DrvInfo(FileSystemManager parent, string name, string path, DriveType type, bool isReady, bool loadDirectories = false, bool loadFiles = false) : base(parent, name, path, loadDirectories, loadFiles)
        {  
            Type = type;
            IsReady = isReady;
        }

        public DriveType Type { get; private set; }

        public override bool IsReady { get; internal set; }
        
        public override bool Exists
        {
            get
            {
                return System.IO.Directory.Exists(Path);
            }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }
        
        public override bool IsFilesLoaded { get; internal set; }
        
        public override bool IsDrivesLoaded { get; internal set; }

        public override bool HasDirectories
        {
            get
            {
                if (!IsDirectoriesLoaded)
                    LoadDirectories();
                return Directories.Length > 0;
            }
        }

        public override bool HasFiles
        {
            get
            {
                if(!IsFilesLoaded)
                    LoadFiles();
                return Files.Length > 0;
            }
        }

        public override void LoadDrives()
        {
            
        }

        public override void LoadFiles()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    FileInfo[] files = dirInfo.GetFiles(String.IsNullOrWhiteSpace(Parent.FileSearchPattern) != true ? Parent.FileSearchPattern : "*.*", SearchOption.TopDirectoryOnly);

                    List<FiInfo> list = new List<FiInfo>();
                    for (int i = 0; i < files.Length; i++)
                    {
                        if (!files[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = files[i].Name;
                            string path = files[i].FullName;
                            list.Add(new FiInfo(Parent, name, path, files[i].Length));
                        }
                    }
                    Files = list.ToArray();
                }
                else
                {
                    Files = new FiInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsFilesLoaded = true;
            }
        }

        public override void LoadDirectories()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    DirectoryInfo[] dirs = dirInfo.GetDirectories("*.*", SearchOption.TopDirectoryOnly);

                    List<FolderInfo> list = new List<FolderInfo>();
                    for (int i = 0; i < dirs.Length; i++)
                    {
                        if (!dirs[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = dirs[i].Name;
                            string path = dirs[i].FullName;
                            list.Add(new FolderInfo(Parent, name, path, false));
                        }
                    }
                    Directories = list.ToArray();
                }
                else
                {
                    Directories = new FolderInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsDirectoriesLoaded = true;
            }
        }
    }

    public class FolderInfo : FileSystemInfo
    {
        internal FolderInfo(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false): base(parent, name, path, loadDirectories, loadFiles)
        {
        }

        public override bool Exists
        {
            get 
            {
                return Directory.Exists(Path);
            }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }
        
        public override bool IsDrivesLoaded { get; internal set; }

        public override bool HasDirectories
        {
            get
            {
                if (!IsDirectoriesLoaded)
                    LoadDirectories();
                return Directories.Length > 0;
            }
        }

        public override bool HasFiles
        {
            get
            {
                if (!IsFilesLoaded)
                    LoadFiles();
                return Files.Length > 0;
            }
        }

        public override void LoadDrives()
        {
        }

        public override void LoadFiles()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    FileInfo[] files = dirInfo.GetFiles(String.IsNullOrWhiteSpace(Parent.FileSearchPattern) != true ? Parent.FileSearchPattern : "*.*", SearchOption.TopDirectoryOnly);

                    List<FiInfo> list = new List<FiInfo>();
                    for (int i = 0; i < files.Length; i++)
                    {
                        if (!files[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = files[i].Name;
                            string path = files[i].FullName;
                            list.Add(new FiInfo(Parent, name, path, files[i].Length));
                        }
                    }
                    Files = list.ToArray();
                }
                else
                {
                    Files = new FiInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsFilesLoaded = true;
            }
        }

        public override void LoadDirectories()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    DirectoryInfo[] dirs = dirInfo.GetDirectories("*.*", SearchOption.TopDirectoryOnly);

                    List<FolderInfo> list = new List<FolderInfo>();
                    for (int i = 0; i < dirs.Length; i++)
                    {
                        if (!dirs[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = dirs[i].Name;
                            string path = dirs[i].FullName;
                            list.Add(new FolderInfo(Parent, name, path, false));
                        }
                    }
                    Directories = list.ToArray();
                }
                else
                {
                    Directories = new FolderInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsDirectoriesLoaded = true;
            }
        }
    }

    public class FiInfo : FileSystemInfo
    {
        internal FiInfo(FileSystemManager parent, string name, string path, long lenght) : base(parent, name, path)
        {
            Size = lenght;
            TryGetAssociatedIcon();
        }

        public override bool Exists
        {
            get 
            {
                return File.Exists(Path);
            }
        }

        public string ValidatedSize
        {
            get 
            {
                return Size.ToString() + " Bytes";
            }
        }

        public long Size { get; private set; }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }
       
        public override bool IsDrivesLoaded { get; internal set; }

        public override void LoadDrives()
        {
        }

        public override void LoadFiles()
        {
        }

        public override void LoadDirectories()
        {
        }
         
        private void TryGetAssociatedIcon()
        {
            //Win32.SHFILEINFO shinfoSmall = new Win32.SHFILEINFO();
            //Win32.SHFILEINFO shinfoLarge = new Win32.SHFILEINFO();

            //IntPtr hImgSmall = Win32.SHGetFileInfo(Path, 0, ref shinfoSmall, (int)Marshal.SizeOf(shinfoSmall), Convert.ToInt32(Win32.SHGFI_ICON | Win32.SHGFI_SMALLICON));
            //IntPtr hImgLarge = Win32.SHGetFileInfo(Path, 0, ref shinfoLarge, (int)Marshal.SizeOf(shinfoLarge), Convert.ToInt32(Win32.SHGFI_ICON | Win32.SHGFI_LARGEICON));
            //if (hImgSmall.ToInt32() > 0 && hImgLarge.ToInt32() > 0 &&
            //    shinfoSmall.hIcon.ToInt32() > 0 && shinfoLarge.hIcon.ToInt32() > 0)
            //{
            //    DisplayImageSmall = System.Drawing.Icon.FromHandle(shinfoSmall.hIcon);
            //    DisplayImageLarge = System.Drawing.Icon.FromHandle(shinfoLarge.hIcon);
            //    if (null != DisplayImageSmall && null != DisplayImageLarge)
            //        SupportsDisplayImage = true;            
            //}
        }
    }

    internal class SpecialFolderRoot : FileSystemInfo
    {
        internal SpecialFolderRoot(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false): base(parent, name, path, loadDirectories, loadFiles)
        { 
        }

        public override bool Exists
        {
            get
            {
                return true;
            }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }

        public override bool IsDrivesLoaded { get; internal set; }

        public override bool HasDirectories
        {
            get
            {
                return true;
            }
        }

        public override bool HasFiles
        {
            get
            {
                return false;
            }
        }

        public override void LoadDrives()
        {
        }

        public override void LoadFiles()
        {
        }

        public override void LoadDirectories()
        {
            try
            {
                List<FolderInfo> list = new List<FolderInfo>();
                string[] folderNames = System.Enum.GetNames(typeof(Environment.SpecialFolder));
                foreach (string folderName in folderNames)
                {
                    Environment.SpecialFolder member = (Environment.SpecialFolder)System.Enum.Parse(typeof(Environment.SpecialFolder), folderName, true);
                    string folderPath = Environment.GetFolderPath(member);
                    FolderInfo fi = new FolderInfo(Parent, member.ToString(), folderPath);
                    list.Add(fi);
                }
                Directories = list.ToArray();
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsDirectoriesLoaded = true;
            }
        }
    }

    internal class SpecialFolderInfo : FileSystemInfo
    {
        internal SpecialFolderInfo(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false) : base(parent, name, path, loadDirectories, loadFiles)
        { 
        }

        public override bool Exists
        {
            get 
            {
                return Directory.Exists(Path);
            }
        }

        public override bool IsDirectoriesLoaded { get; internal set; }

        public override bool IsFilesLoaded { get; internal set; }
        
        public override bool IsDrivesLoaded { get; internal set; }

        public override bool HasDirectories
        {
            get
            {
                if (!IsDirectoriesLoaded)
                    LoadDirectories();
                return Directories.Length > 0;
            }
        }

        public override bool HasFiles
        {
            get
            {
                if (!IsFilesLoaded)
                    LoadFiles();
                return Files.Length > 0;
            }
        }

        public override void LoadDrives()
        { 
        }

        public override void LoadFiles()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    FileInfo[] files = dirInfo.GetFiles(String.IsNullOrWhiteSpace(Parent.FileSearchPattern) != true ? Parent.FileSearchPattern : "*.*", SearchOption.TopDirectoryOnly);

                    List<FiInfo> list = new List<FiInfo>();
                    for (int i = 0; i < files.Length; i++)
                    {
                        if (!files[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = files[i].Name;
                            string path = files[i].FullName;
                            list.Add(new FiInfo(Parent, name, path, files[i].Length));
                        }
                    }
                    Files = list.ToArray();
                }
                else
                {
                    Files = new FiInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsFilesLoaded = true;
            }
        }

        public override void LoadDirectories()
        {
            try
            {
                if (Exists)
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(Path);
                    DirectoryInfo[] dirs = dirInfo.GetDirectories("*.*", SearchOption.TopDirectoryOnly);

                    List<FolderInfo> list = new List<FolderInfo>();
                    for (int i = 0; i < dirs.Length; i++)
                    {
                        if (!dirs[i].Attributes.HasFlag(FileAttributes.System))
                        {
                            string name = dirs[i].Name;
                            string path = dirs[i].FullName;
                            list.Add(new FolderInfo(Parent, name, path, false));
                        }
                    }
                    Directories = list.ToArray();
                }
                else
                {
                    Directories = new FolderInfo[0];
                }
            }
            catch (UnauthorizedAccessException)
            {
                HasErrors = true;
            }
            catch (Exception)
            {
                throw;
            }
            finally
            {
                IsDirectoriesLoaded = true;
            }
        }
    }

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

            //List<SpecialFolderInfo> list = new List<SpecialFolderInfo>();
            //string[] folderNames = System.Enum.GetNames(typeof(Environment.SpecialFolder));
            //foreach (string folderName in folderNames)
            //{
            //    Environment.SpecialFolder member = (Environment.SpecialFolder)System.Enum.Parse(typeof(Environment.SpecialFolder), folderName, true);
            //    string folderPath = Environment.GetFolderPath(member);
            //    SpecialFolderInfo folderInfo = new SpecialFolderInfo(this, folderName, folderPath, load);
            //    list.Add(folderInfo);
            //}
            //return list.ToArray();
        }

        

        private static string CalculateVolumeLabel(DriveInfo driveInfo)
        {
            if (driveInfo.DriveType == DriveType.Network)
            {
                if (!driveInfo.IsReady)
                    return "";

                string rootDirectory = driveInfo.RootDirectory.Root.Name;
                string result = ResolveShareName(driveInfo.RootDirectory.Root.Name);

                if (rootDirectory.EndsWith("\\", StringComparison.InvariantCultureIgnoreCase))
                    rootDirectory = rootDirectory.Substring(0, rootDirectory.Length - 1);

                if (result.StartsWith("(" + rootDirectory + ")", StringComparison.InvariantCultureIgnoreCase))
                {
                    result = result.Substring(("(" + rootDirectory + ")").Length);
                    result = result.Trim();
                }
                if(!String.IsNullOrWhiteSpace(result))
                    result = " (" + result + ")";

                return result;
            }
            else if (driveInfo.DriveType == DriveType.CDRom)
            {
                if (driveInfo.IsReady)
                    return " (" + driveInfo.VolumeLabel + ")";
                else
                    return " (CD-Laufwerk)";
            }
            else
            {
                if (!driveInfo.IsReady)
                    return "";
                return " (" + driveInfo.VolumeLabel + ")";
            }
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

        //internal bool DriveHasDirectoriesOrFiles(DriveInfo drive)
        //{
        //    if (!drive.IsReady)
        //        return false;

        //    DirectoryInfo dirInfo = new DirectoryInfo(drive.RootDirectory.Name);
        //    dirInfo.GetDirectories();
        //}
    }
}
