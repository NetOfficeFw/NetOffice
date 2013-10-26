using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Represents an Item in local filesystem
    /// </summary>
    public abstract class FileSystemInfo
    {
        #region Ctor
        
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

        #endregion

        #region Properties

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
        /// Returns the instance has already loaded folders from the filesystem
        /// </summary>
        public abstract bool IsDirectoriesLoaded { get; internal set; }

        /// <summary>
        /// Returns the instance has already loaded files from the filesystem
        /// </summary>
        public abstract bool IsFilesLoaded { get; internal set; }

        /// <summary>
        /// Returns the instance has already loaded drives from the filesystem
        /// </summary>
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
        /// Drives if exists
        /// </summary>
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

        #endregion

        #region Methods

        /// <summary>
        /// Load or reload files from filesystem
        /// </summary>
        public abstract void LoadFiles();

        /// <summary>
        /// Load or reload folders from filesystem
        /// </summary>
        public abstract void LoadDirectories();

        /// <summary>
        ///  Load or reload drives from filesystem
        /// </summary>
        public abstract void LoadDrives();

        #endregion
    }
}
