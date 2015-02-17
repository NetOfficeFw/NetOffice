using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using System.Text;
using ICSharpCode.SharpZipLib;
using ICSharpCode.SharpZipLib.Core;
using ICSharpCode.SharpZipLib.Zip;

namespace NetOffice.DeveloperToolbox.Translation
{
    /// <summary>
    /// Configurable language settings
    /// </summary>
    public class ToolLanguage : INotifyPropertyChanged
    {
        #region Fields

        private string _name = String.Empty;
        private string _nameGlobal = String.Empty;
        private string _authorName = String.Empty;
        private string _authorMail = String.Empty;
        private string _authorSite = String.Empty;
        private int _lcid = 0;
        private ToolLanguages _parent;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class as existing language
        /// </summary>
        /// <param name="parent">parent collection</param>
        /// <param name="nameGlobal">international name</param>
        /// <param name="lcid">language id</param>
        internal ToolLanguage(ToolLanguages parent, string nameGlobal, int lcid)
        {
            _parent = parent;
            _nameGlobal = nameGlobal;
            _lcid = lcid;
        }

        /// <summary>
        /// Creates an instance of the class as new language
        /// </summary>
        /// <param name="parent">parent collection</param>
        /// <param name="template">role model</param>
        internal ToolLanguage(ToolLanguages parent, ToolLanguage template)
        {
            _parent = parent;
            _name = "New Language";
            _nameGlobal = "New Language";
            _lcid = 0;
            Initialize();

            foreach (var item in template.Application.Components[0].ControlRessources)
                Application.Components[0].ControlRessources[item.Value].Value2 = item.Value2;

            foreach (var item in template.Application.Components[1].ControlRessources)
                Application.Components[1].ControlRessources[item.Value].Value2 = item.Value2;

            for (int i = 0; i < template.Components.Count; i++)
            {
                LocalizableCompoment templateComponent = template.Components[i];
                LocalizableCompoment ownComponent = Components[i];
                foreach (var item in templateComponent.ControlRessources)
                    ownComponent.ControlRessources[item.Value].Value2 = item.Value2;
            }

            IsNew = true;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Language is new and unsaved
        /// </summary>
        internal bool IsNew { get; set; }

        /// <summary>
        /// Language contains unsaved changed
        /// </summary>
        internal bool IsDirty { get; set; }

        /// <summary>
        /// The user level name
        /// </summary>
        public virtual string DisplayName
        {
            get
            {
                if (String.IsNullOrWhiteSpace(Name))
                    return "<Empty>";
                if (Name.Equals(NameGlobal, StringComparison.InvariantCultureIgnoreCase))
                    return Name;
                else
                    return String.Format("{0} ({1})", NameGlobal, Name);
            }
        }

        /// <summary>
        /// The logical name
        /// </summary>
        public virtual string Name
        {
            get 
            {
                return _name;
            }
            set
            {
                _name = value;
                RaiseNotifyPropertyChanged("Name");
            }
        }

        /// <summary>
        /// The international name
        /// </summary>
        public virtual string NameGlobal
        {
            get
            {
                return _nameGlobal;
            }
            set
            {
                _nameGlobal = value;
                RaiseNotifyPropertyChanged("NameGlobal");
            }
        }

        /// <summary>
        /// Name of the author from the language
        /// </summary>
        public virtual string Author
        {
            get
            {
                return _authorName;
            }
            set
            {
                _authorName = value;
                RaiseNotifyPropertyChanged("Author");
            }
        }

        /// <summary>
        /// Mail adress from the author
        /// </summary>
        public virtual string AuthorMail
        {
            get
            {
                return _authorMail;
            }
            set
            {
                _authorMail = value;
                RaiseNotifyPropertyChanged("AuthorMail");
            }
        }

        /// <summary>
        /// Website from the author
        /// </summary>
        public virtual string AuthorSite
        {
            get
            {
                return _authorSite;
            }
            set
            {
                _authorSite = value;
                RaiseNotifyPropertyChanged("AuthorSite");
            }
        }

        /// <summary>
        /// Language ID
        /// </summary>
        public virtual int LCID
        {
            get
            {
                return _lcid;
            }
            set
            {
                if(false == IsNew && value <= 1000 || value >= 32000)
                    throw new ArgumentOutOfRangeException("Valid range is 1000-32000");
                foreach (var item in _parent)
                    if (item != this && item.LCID == value)
                        throw new ArgumentException("Duplicate LCID");
                _lcid = value;
                RaiseNotifyPropertyChanged("LCID");
            }
        }

        /// <summary>
        /// Application Root Components
        /// </summary>
        internal ToolLanguageApplication Application { get; private set; }

        /// <summary>
        /// Application Plugin Components
        /// </summary>
        internal LocalizableCompoments Components { get; private set; }

        #endregion

        #region Methods

        /// <summary>
        /// Returns the language contains all necessary informations to save them to filesystem
        /// </summary>
        /// <returns></returns>
        internal bool IsValid()
        {
            if (String.IsNullOrWhiteSpace(_nameGlobal))
                return false;
            if (_lcid <= 1000 || _lcid >= 32000)
                return false;

            return true;
        }

        /// <summary>
        /// Get localizable values (Not implemented)
        /// </summary>
        /// <param name="componentName">name of the component</param>
        /// <returns>items</returns>
        internal virtual ItemCollection GetValues(string componentName)
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Load language from filesystem
        /// </summary>
        /// <param name="fileName">full qualified file name</param>
        internal virtual void Load(string fileName)
        {
            using (ZipFile file = new ZipFile(fileName))
            {
                foreach (ZipEntry item in file)
                {
                    using (Stream stream = file.GetInputStream(item))
                    {
                        StreamReader reader = new StreamReader(stream);
                        string content = reader.ReadToEnd();
                        XDocument document = XDocument.Parse(content);
                        XElement root = document.Root as XElement;
                        switch (root.Name.LocalName)
                        {
                            case "NetOffice.DeveloperToolbox.Translation.ToolLanguage":
                                ReadLanguageSummary(root);
                                break;
                            case "NetOffice.DeveloperToolbox.Translation.Component":
                                ReadApplicationComponent(item.Name, root);
                                break;
                            default:
                                break;
                        }
                        reader.Close();
                        reader.Dispose();
                    }
                }
            }
        }

        /// <summary>
        /// Save language to file
        /// </summary>
        internal virtual void Save()
        {
            if (!IsValid())
                throw new InvalidOperationException("Invalid settings");

            if (!Directory.Exists(ToolLanguages.DirectoryPath))
                Directory.CreateDirectory(ToolLanguages.DirectoryPath);
            string targetFilePath = Path.Combine(ToolLanguages.DirectoryPath, LCID + ToolLanguages.Extension);
            
            string tempPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempPath);

            XDocument indexDocument = new XDocument();
            XElement rootNode = new XElement("NetOffice.DeveloperToolbox.Translation.ToolLanguage");
            rootNode.Add(new XElement("NameGlobal", System.Xml.XmlConvert.EncodeName(NameGlobal)));
            rootNode.Add(new XElement("NameLocal", System.Xml.XmlConvert.EncodeName(Name)));
            rootNode.Add(new XElement("LCID", System.Xml.XmlConvert.EncodeName(LCID.ToString())));
            rootNode.Add(new XElement("AuthorName", System.Xml.XmlConvert.EncodeName(Author)));
            rootNode.Add(new XElement("AuthorMail", System.Xml.XmlConvert.EncodeName(AuthorMail)));
            rootNode.Add(new XElement("AuthorSite", System.Xml.XmlConvert.EncodeName(AuthorSite)));
            indexDocument.Add(rootNode);
            indexDocument.Save(Path.Combine(tempPath, "Index"));

            foreach (var item in Application.Components)
            {
                XDocument appComponentDocument = new XDocument();
                XElement rootAppComponent = new XElement("NetOffice.DeveloperToolbox.Translation.Component", new XAttribute("IsSystem", true));

                foreach (var subItem in item.ControlRessources)
                    rootAppComponent.Add(new XElement("Pair", new XElement("Name", subItem.Value), new XElement("Value", subItem.Value2)));

                rootAppComponent.Save(Path.Combine(tempPath, item.Value));
            }

            foreach (var item in Components)
            {
                XDocument appComponentDocument = new XDocument();
                XElement rootAppComponent = new XElement("NetOffice.DeveloperToolbox.Translation.Component", new XAttribute("IsSystem", false));

                foreach (var subItem in item.ControlRessources)
                    rootAppComponent.Add(new XElement("Pair", new XElement("Name", subItem.Value), new XElement("Value", System.Xml.XmlConvert.EncodeName(subItem.Value2))));
                         rootAppComponent.Save(Path.Combine(tempPath, item.Value));
            }

            FileStream fsOut = File.Create(targetFilePath);
            ZipOutputStream zipStream = new ZipOutputStream(fsOut);
            int folderOffset = tempPath.Length + (tempPath.EndsWith("\\") ? 0 : 1);
            CompressFolder(tempPath, zipStream, folderOffset);
            zipStream.IsStreamOwner = true;
            zipStream.Close();

            Directory.Delete(tempPath, true);

            _parent.ValidateFiles();

            IsNew = false;
            IsDirty = false;
        }

        /// <summary>
        /// Initialize the language
        /// </summary>
        internal virtual void Initialize()
        {
            Application = new ToolLanguageApplication(this);
            Components = new LocalizableCompoments();

            string space = " - ";
            LocalizableCompoment comp = null;
            comp =  new LocalizableCompoment(this, "Welcome", typeof(ToolboxControls.Welcome.WelcomeControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.Welcome.WelcomeControl).FullName, "Welcome" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);                
            }

            comp = new LocalizableCompoment(this, "Office Compatibility", typeof(ToolboxControls.OfficeCompatibility.OfficeCompatibilityControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.OfficeCompatibility.OfficeCompatibilityControl).FullName, "Office Compatibility" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "Application Observer", typeof(ToolboxControls.ApplicationObserver.ApplicationObserverControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.ApplicationObserver.ApplicationObserverControl).FullName, "Application Observer" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "Registry Editor", typeof(ToolboxControls.RegistryEditor.RegistryEditorControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.RegistryEditor.RegistryEditorControl).FullName, "Registry Editor" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "Addin Guard", typeof(ToolboxControls.AddinGuard.AddinGuardControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.AddinGuard.AddinGuardControl).FullName, "Addin Guard" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "Office UI", typeof(ToolboxControls.OfficeUI.OfficeUIControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.OfficeUI.OfficeUIControl).FullName, "Office UI" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "Outlook Security", typeof(ToolboxControls.OutlookSecurity.OutlookSecurityControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.OutlookSecurity.OutlookSecurityControl).FullName, "Outlook Security" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }
            comp = new LocalizableCompoment(this, "Project Wizard", typeof(ToolboxControls.ProjectWizard.ProjectWizardControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.ProjectWizard.ProjectWizardControl).FullName, "Project Wizard" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }

            comp = new LocalizableCompoment(this, "About", typeof(ToolboxControls.About.AboutControl));
            Components.Add(comp);
            foreach (var item in (comp.Design as ILocalizationDesign).Childs)
            {
                comp = new LocalizableCompoment(this, typeof(ToolboxControls.About.AboutControl).FullName, "About" + space + item.NameLocalization, item.TypeLocalization);
                Components.Add(comp);
            }
        }

        private static void CompressFolder(string path, ZipOutputStream zipStream, int folderOffset)
        {
            string[] files = Directory.GetFiles(path);

            foreach (string filename in files)
            {
                FileInfo fi = new FileInfo(filename);
                string entryName = filename.Substring(folderOffset);
                entryName = ZipEntry.CleanName(entryName);
                ZipEntry newEntry = new ZipEntry(entryName);
                newEntry.DateTime = fi.LastWriteTime; 
                newEntry.Size = fi.Length;
                zipStream.PutNextEntry(newEntry);
                byte[] buffer = new byte[4096];
                using (FileStream streamReader = File.OpenRead(filename))
                {
                    StreamUtils.Copy(streamReader, zipStream, buffer);
                }
                zipStream.CloseEntry();
            }
        }

        private void ReadLanguageSummary(XElement element)
        {
            NameGlobal = System.Xml.XmlConvert.DecodeName(element.Element("NameGlobal").Value);
            Name = System.Xml.XmlConvert.DecodeName(element.Element("NameLocal").Value);
            LCID = Convert.ToInt32(System.Xml.XmlConvert.DecodeName(element.Element("LCID").Value));
            Author = System.Xml.XmlConvert.DecodeName(element.Element("AuthorName").Value);
            AuthorMail = System.Xml.XmlConvert.DecodeName(element.Element("AuthorMail").Value);
            AuthorSite = System.Xml.XmlConvert.DecodeName(element.Element("AuthorSite").Value);
        }

        private void ReadApplicationComponent(string componentName, XElement element)
        {
            LocalizableCompoment component = null;
            bool isSystem = Convert.ToBoolean(element.Attribute("IsSystem").Value);
            if(isSystem)
                component = Application.Components[componentName];
            else
                component = Components[componentName];

            foreach (var item in component.ControlRessources)
            {
                XElement el = element.Elements("Pair").Where(e => e.Element("Name").Value == item.Value).FirstOrDefault();
                if (null != el)
                    item.Value2 = System.Xml.XmlConvert.DecodeName(el.Element("Value").Value);
            }
        }

        #endregion

        #region INotifyPropertyChanged

        public event PropertyChangedEventHandler  PropertyChanged;

        private void RaiseNotifyPropertyChanged(string propertyName)
        {
            if (null != PropertyChanged)
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a System.String that represents the instance
        /// </summary>
        /// <returns></returns>
        public override string ToString()
        {
            return String.Format("ToolLanguage {0}", String.IsNullOrWhiteSpace(NameGlobal) ? "<NoName>" : NameGlobal);
        }

        #endregion
    }
}
