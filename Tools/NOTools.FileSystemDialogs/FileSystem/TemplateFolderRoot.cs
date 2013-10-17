using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// TemplateFolders Collection Item
    /// </summary>
    internal class TemplateFolderRoot : FileSystemInfo
    {
        #region Ctor
        
        internal TemplateFolderRoot(FileSystemManager parent, string name, string path, TemplateFolderDescription[] description, bool loadDirectories = false, bool loadFiles = false)  : base(parent, name, path, loadDirectories, loadFiles)
        {
            if (null != description)
                Description = description;
            else
                Description = new TemplateFolderDescription[0];
        }

        #endregion

        #region Properties

        /// <summary>
        /// Source Definition
        /// </summary>
        internal TemplateFolderDescription[] Description { get; private set; }

        #endregion

        #region Overrides

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
                return Description.Length > 0;
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
                foreach (TemplateFolderDescription item in Description)
                {
                    string folderPath = item.Path;
                    FolderInfo fi = new FolderInfo(Parent, item.DisplayName, folderPath);
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
     
        #endregion
    }
}
