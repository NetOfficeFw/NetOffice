using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// SpecialFolders Collection Item
    /// </summary>
    internal class SpecialFolderRoot : FileSystemInfo
    {
        #region Ctor

        internal SpecialFolderRoot(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false) : base(parent, name, path, loadDirectories, loadFiles)
        {
        }

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

        #endregion
    }
}
