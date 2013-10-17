using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.FileSystemDialogs
{
    /// <summary>
    /// Main handler to the local filesystem
    /// </summary>
    internal class FileSystemManager
    {
        #region Properties

        /// <summary>
        /// Additional search pattern
        /// </summary>
        internal string FileSearchPattern { get; set; }

        #endregion

        #region Methods
        
        /// <summary>
        /// Returns Desktop folder item
        /// </summary>
        /// <param name="load">load info immediately</param>
        /// <returns>Desktop folder item</returns>
        internal FolderInfo GetDesktop(bool load = false)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string name = Path.GetFileName(folderPath);
            FolderInfo folderInfo = new FolderInfo(this, name, folderPath, load);
            return folderInfo;
        }

        /// <summary>
        /// Returns MyComputer item
        /// </summary>
        /// <returns>MyComputer item</returns>
        internal DrvInfoRoot GetMyComputer()
        {
            DrvInfoRoot driveRoot = new DrvInfoRoot(this, "MyComputer","");
            return driveRoot;
        }

        /// <summary>
        /// Returns MyDocuments folder item
        /// </summary>
        /// <param name="load">load info immediately</param>
        /// <returns>MyDocuments folder item</returns>
        internal FolderInfo GetMyDocuments(bool load = false)
        {
            string folderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string name = Path.GetFileName(folderPath);
            FolderInfo folderInfo = new FolderInfo(this, name, folderPath, load);
            return folderInfo;
        }

        /// <summary>
        /// Returns SpecialFolders item
        /// </summary>
        /// <returns>SpecialFolders item</returns>
        internal SpecialFolderRoot GetSpecialFolders()
        {
            SpecialFolderRoot driveRoot = new SpecialFolderRoot(this, "SpecialFolders", "");
            return driveRoot;
        }

        /// <summary>
        /// Returns TemplateFolders item
        /// </summary>
        /// <returns>TemplateFolders item</returns>
        internal TemplateFolderRoot GetTemplateFolders(TemplateFolderDescription[] description)
        {
            TemplateFolderRoot folderRooth = new TemplateFolderRoot(this, "TemplateFolders", "", description);
            return folderRooth;
        }

        #endregion

    }
}
