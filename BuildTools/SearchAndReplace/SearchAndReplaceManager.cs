using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOBuildTools.SearchAndReplace
{
    /// Progress log action handler
    /// </summary>
    /// <param name="message">log action message</param>
    public delegate void LogAction(string message);

    /// <summary>
    /// Search and replace logic for text files
    /// </summary>
    internal static class SearchAndReplaceManager
    {
        /// <summary>
        /// Read all files in a directory and replace arg search with arg replace in file content(s)
        /// </summary>
        /// <param name="directoryName">target root directory</param>
        /// <param name="fileFilter">exclude filter as file extension</param>
        /// <param name="search">search expression</param>
        /// <param name="replace">replace value</param>
        /// <param name="func">log handler</param>
        public static void SearchAndReplace(string directoryName, string fileFilter, string search, string replace, LogAction func)
        {
            if (!Directory.Exists(directoryName))
                throw new DirectoryNotFoundException(directoryName);
            if (null == func || String.IsNullOrWhiteSpace(replace) || String.IsNullOrWhiteSpace(search) || String.IsNullOrWhiteSpace(fileFilter) || String.IsNullOrWhiteSpace(directoryName))
                throw new ArgumentNullException();

            string[] filterArray = BuildFilterArray(fileFilter);
            string[] searchArray = BuildSearchArray(search);
            string[] replaceArray = BuildReplaceArray(replace);
            if (searchArray.Length != replaceArray.Length)
                throw new FormatException("Search and Repleace terms count must equal");
            SearchAndReplace(directoryName, filterArray, searchArray, replaceArray, func);
        }

        private static string[] BuildFilterArray(string fileFilter)
        {
            List<string> list = new List<string>();
            if (!String.IsNullOrWhiteSpace(fileFilter))
            {
                string[] tempArray = fileFilter.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in tempArray)
                    list.Add(item.Trim());
                if (list.Count == 0)
                    list.Add("*.*");
            }
            else
                list.Add("*.*");
            return list.ToArray();
        }

        private static string[] BuildSearchArray(string search)
        {
            List<string> list = new List<string>();
            if (!String.IsNullOrWhiteSpace(search))
            {
                string[] tempArray = search.Split(new string[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var item in tempArray)
                    list.Add(item.Trim());
                if(list.Count == 0)
                    throw new FormatException("Search term can't be empty.");
            }
            else
                throw new FormatException("Search term can't be empty.");
            return list.ToArray();
        }

        private static string[] BuildReplaceArray(string replace)
        {
            List<string> list = new List<string>();
            if (!String.IsNullOrWhiteSpace(replace))
            {
                string[] tempArray = replace.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                foreach (var item in tempArray)
                    list.Add(item.Trim());
                if (list.Count == 0)
                    throw new FormatException("Replace term can't be empty.");
            }
            else
                throw new FormatException("Replace term can't be empty.");
            return list.ToArray();
        }

        private static bool FilterPassed(string file, string[] fileFilter)
        {
            foreach (var item in fileFilter)
            {
                string end = item;
                string extenstion = Path.GetExtension(file);
                if (extenstion.StartsWith("."))
                    extenstion = extenstion.Substring(1);
                int pos = item.IndexOf(".");
                if (pos > -1)
                    end = item.Substring(pos + 1);
                if (end.Equals(extenstion, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        private static bool DoSearchAndReplace(ref string fileContent, string[] searchArray, string[] replaceArray, LogAction func)
        {
            bool oneOrMoreReplaced = false;
            for (int i = 0; i < searchArray.Length; i++)
            {
                string search = searchArray[i];
                string replace = replaceArray[i];
                int cnt = 0;
                while (fileContent.IndexOf(search) > -1)
                {
                    fileContent = fileContent.Replace(search, replace);
                    cnt++;
                }
                if (cnt > 0)
                {
                    oneOrMoreReplaced = true;
                    func(String.Format("{0} entries of {1} replaced", cnt, search));
                }

            }

            return oneOrMoreReplaced;
        }

        private static bool WriteFile(string file, string fileContent, LogAction func)
        {
            try 
	        {
                File.WriteAllText(file, fileContent);
                return true;
	        }
            catch (Exception exception)
            {
                func("File write error." + exception.ToString());
                return false;
            }
        }

        private static bool ReverseMoveBackupFile(string backupFile, string originFile)
        {
            try
            {
                File.Move(backupFile, originFile);
                return true;
            }
            catch 
            {
                return false;
            }
        }

        private static string CopyFileBackup(string file, LogAction func)
        {
            try
            {
                string fileName = System.IO.Path.GetFileName(file);
                string newFile = System.IO.Path.Combine(Application.StartupPath, "BackupSearchAndReplace", fileName);

                if(!Directory.Exists( System.IO.Path.Combine(Application.StartupPath, "BackupSearchAndReplace")))
                    Directory.CreateDirectory(System.IO.Path.Combine(Application.StartupPath, "BackupSearchAndReplace"));

                File.Copy(file, newFile);
                return newFile;
            }
            catch (Exception exception)
            {
                func("File copy error." + exception.ToString());
                return null;
            }
        }

        private static bool FileDelete(string file, LogAction func)
        {
            try
            {
                File.Delete(file);
                return true;
            }
            catch (Exception exception)
            {
                func("File delete error." + exception.ToString());
                return false;
            }
        }

        private static bool ReadFile(string file, ref string fileContent, LogAction func)
        {
            try
            {
                fileContent = File.ReadAllText(file, Encoding.UTF8);
                return true;
            }
            catch (Exception exception)
            {
                func("File reading error." + exception.ToString());
                return false;
            }
        }

        private static void DeleteBackupFile(string fileName)
        {
            File.Delete(fileName);
        }

        private static void SearchAndReplace(string directoryName, string[] fileFilter, string[] search, string[] replace, LogAction func)
        {
            int i = 0;
            func("Search and Replace is started");
            foreach (var item in Directory.GetFiles(directoryName, "*.*", SearchOption.AllDirectories))
            {
                i++;
                if(i.ToString().EndsWith("00"))
                    func("");

                bool filterPassed = FilterPassed(item, fileFilter);
                if (!filterPassed)
                    continue;
                string fileContent = string.Empty;
                if (ReadFile(item, ref fileContent, func))
                { 
                    if (DoSearchAndReplace(ref fileContent, search, replace, func))
                    {
                        string newFileName = CopyFileBackup(item, func);
                        if (null != newFileName)
                        { 
                            if(FileDelete(item, func))
                            {
                                if (WriteFile(item, fileContent, func))
                                {
                                    DeleteBackupFile(newFileName);
                                    func("changed: " + item);
                                }
                                else
                                { 
                                     if(ReverseMoveBackupFile(newFileName, item))
                                         func("backup file restored after write error: " + item);
                                     else
                                         func("WARNING: backup file not restored after write error: " + item);
                                }
                            }
                            
                        }
                    }
                }
            }
            func("Search and Replace is complete");
        }
    }
}
