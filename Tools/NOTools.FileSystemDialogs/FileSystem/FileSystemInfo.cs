using System;
using System.Drawing;
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
        public virtual Icon DisplayImageSmall { get; internal set; }

        /// <summary>
        /// Display Image(32x32) (if SupportsDisplayImage is okay)
        /// </summary>
        public virtual Icon DisplayImageLarge { get; internal set; }


        public DrvInfo[] Drives
        {
            get
            {
                //if (!IsDrivesLoaded)
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
                //if (!IsDirectoriesLoaded)
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
                //if (!IsFilesLoaded)
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
}
