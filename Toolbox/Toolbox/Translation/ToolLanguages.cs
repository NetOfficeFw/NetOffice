using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Contains all supported languages
    /// </summary>
    public class ToolLanguages : BindingList<ToolLanguage>, IDisposable
    {
        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        internal ToolLanguages()
        {
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.English));
            this[0].Initialize();
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.German));
            this[1].Initialize();
        }

        /// <summary>
        /// File extension to load and save languages from file
        /// </summary>
        internal static string Extension
        {
            get 
            {
                return ".lng";
            }
        }

        /// <summary>
        /// File extension pattern to load files
        /// </summary>
        internal static string ExtensionWildCard
        {
            get
            {
                return "*.lng";
            }
        }

        /// <summary>
        /// Directory to load and save languages
        /// </summary>
        internal static string DirectoryPath
        {
            get
            { 
                return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NODevBox\\Languages");
            }
        }

        /// <summary>
        /// Contains a language
        /// </summary>
        /// <param name="lcid">language id</param>
        /// <returns>true if exists otherwise false</returns>
        public bool Contains(int lcid)
        {
            return null != this.Where(l => l.LCID == lcid).FirstOrDefault();        
        }

        /// <summary>
        /// Get application language
        /// </summary>
        /// <param name="lcid">language id</param>
        /// <param name="throwExceptionIfNotFound">throw exception if not found otherwise return null</param>
        /// <returns>language or null</returns>
        public ToolLanguage this[int lcid, bool throwExceptionIfNotFound]
        {
            get 
            {
                var item = this.Where(l => l.LCID == lcid).FirstOrDefault();
                if (null != item)
                    return item;
                else if (throwExceptionIfNotFound)
                    throw new ArgumentOutOfRangeException("lcid");
                else
                    return null;
            }
        }

        /// <summary>
        /// Try to find non-used languages in directory and delete them
        /// </summary>
        internal void ValidateFiles()
        {
            if (!Directory.Exists(ToolLanguages.DirectoryPath))
                return;

            List<string> files = new List<string>();
            foreach (var item in Directory.GetFiles(ToolLanguages.DirectoryPath, ExtensionWildCard, SearchOption.TopDirectoryOnly))
            {
                string itemLCID = Path.GetFileNameWithoutExtension(item);
                int lcid = 0;
                if(!int.TryParse(itemLCID, out lcid))
                    files.Add(item);            
                else if (!this.Contains(lcid))
                    files.Add(item);
            }
            foreach (var item in files)
            {
                try
                {
                    File.Delete(item);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);                    
                }
            }
        }

        /// <summary>
        /// Lookup to copy embed language packages to the file-sytem
        /// </summary>
        internal void InitializeFolder()
        {
            try
            {
                // a little code smell(because hardcoded). todo: perform reflection on assembly for .lng files
                System.Reflection.Assembly assembly =  System.Reflection.Assembly.GetExecutingAssembly();
                ExtractLanguagePackage(assembly, "1049"); 
            }
            catch (Exception exception)
            {
                // no need to confusing the user
                Console.WriteLine(exception);
            }
        }

        /// <summary>
        /// Load all languages from folder
        /// </summary>
        internal void LoadFromFolder()
        {
            InitializeFolder();
            Clear();
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.English));
            this[0].Initialize();
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.German));
            this[1].Initialize();

            if (!Directory.Exists(DirectoryPath))
                return;

            string[] files = Directory.GetFiles(DirectoryPath, ExtensionWildCard, SearchOption.TopDirectoryOnly);
            foreach (var item in files)
            {
                string itemLCID = Path.GetFileNameWithoutExtension(item);
                int lcid = 0;
                int.TryParse(itemLCID, out lcid);
                ToolLanguage language = new ToolLanguage(this, "<Not loaded>", lcid);
                try
                {
                    language.Initialize();
                    language.Load(item);
                    Add(language);
                }
                catch (Exception exception)
                {
                    Console.WriteLine(exception);
                }
            }
        }

        /// <summary>
        /// Extract resource and copy to file system if not exists
        /// </summary>
        /// <param name="assembly">toolbox assembly</param>
        /// <param name="fileName">target file</param>
        private void ExtractLanguagePackage(System.Reflection.Assembly assembly, string fileName)
        {
            string targetFilePath = Path.Combine(DirectoryPath, fileName + Extension);
            if (!File.Exists(targetFilePath))
            {
                Stream stream = assembly.GetManifestResourceStream(assembly.GetName().Name + "." + fileName + Extension);
                if (null != stream)
                {
                    byte[] bytes = new byte[stream.Length];
                    stream.Read(bytes, 0, (int)stream.Length);
                    FileStream fs = new FileStream(targetFilePath, FileMode.Create);
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Close();
                    fs.Dispose();
                    stream.Close();
                    stream.Dispose();
                }
            }
        }

        /// <summary>
        /// Dispose the instance
        /// </summary>
        public void Dispose()
        {
            // unsafed
            foreach (var item in this)
            {
                if (item.IsNew || item.IsDirty)
                {
                    try
                    {
                        if(item.IsValid())
                            item.Save();
                    }
                    catch (Exception exception)
                    {
                        Console.WriteLine(exception);
                    }
                }
            }

            // someone to delete?
            ValidateFiles();
        }
    }
}
