using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    public class FolderInfo : FileSystemInfo
    {
        internal FolderInfo(FileSystemManager parent, string name, string path, bool loadDirectories = false, bool loadFiles = false)
            : base(parent, name, path, loadDirectories, loadFiles)
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
}
