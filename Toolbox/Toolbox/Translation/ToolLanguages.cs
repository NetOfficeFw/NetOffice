using System;
using System.IO;
using System.ComponentModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib;

namespace NetOffice.DeveloperToolbox.Translation
{
    public class ToolLanguages : BindingList<ToolLanguage>, IDisposable
    {
        internal ToolLanguages()
        {
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.English));
            this[0].Initialize();
            Add(new ToolDefaultLanguage(this, ToolDefaultLanguageName.German));
            this[1].Initialize();
        }

        internal static string Extension
        {
            get 
            {
                return ".lng";
            }
        }

        internal static string ExtensionWildCard
        {
            get
            {
                return "*.lng";
            }
        }

        internal static string DirectoryPath
        {
            get
            {
                //#if DEBUG                 
                //    return Path.Combine(Program.SubFolder, "Languages");
                //#else
                    return Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "NODevBox\\Languages");
//                #endif
            }
        }

        public bool Contains(int lcid)
        {
            return null != this.Where(l => l.LCID == lcid).FirstOrDefault();        
        }

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

        internal void LoadFromFolder()
        {
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
