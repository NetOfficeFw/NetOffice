using System;
using System.Reflection;
using System.IO;
using System.Xml;
using System.Collections.Generic;
using System.Text;

namespace NetOffice.OfficeApi.Tools.Dialogs
{
    /// <summary>
    /// Represents a set of localized values for NetOffice Tools based default dialogs
    /// </summary>
    public class DialogLocalization : IEnumerable<KeyValuePair<string,string>>
    {
        #region Fields

        private Dictionary<string, string> _list;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="lcid">lanuage id</param>
        /// <param name="values">localized values</param>
        internal DialogLocalization(int lcid, IEnumerable<KeyValuePair<string, string>> values)
        {
            if (lcid <= 0)
                throw new ArgumentOutOfRangeException("lcid");
            if (null == values)
                throw new ArgumentNullException("values");

            _list = new Dictionary<string, string>();
            LCID = lcid;
            foreach (KeyValuePair<string,string> item in values)
                _list.Add(item.Key, item.Value);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Language Id
        /// </summary>
        public int LCID { get; private set; }

        /// <summary>
        /// Get localized value by name
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="faulty">default value if not found</param>
        /// <returns>localized value</returns>
        public string this[string name, string faulty]
        {
            get
            {
                if (_list.ContainsKey(name))
                    return _list[name];
                else
                    return faulty;
            }
        }

        /// <summary>
        /// Get or set localized value
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <returns>localized value</returns>
        public string this[string name]
        {
            get
            {
                return _list[name];
            }
            set
            {
                _list[name] = value;
            }
        }

        /// <summary>
        /// Get or set localized value
        /// </summary>
        /// <param name="name">name of the value</param>
        /// <param name="throwExceptionIfNotFound">throw exception if not found otherwise return null</param>
        /// <returns>localized value of null if throwExceptionIfNotFound is set</returns>
        public string this[string name, bool throwExceptionIfNotFound]
        {
            get
            {
                if (throwExceptionIfNotFound)
                    return _list[name];
                else
                {
                    if (_list.ContainsKey(name))
                        return _list[name];
                    else
                        return null;
                }
            }
            set
            {
                if (throwExceptionIfNotFound && false == _list.ContainsKey(name))
                    throw new ArgumentOutOfRangeException("name");
                _list[name] = value;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add a new localized value
        /// </summary>
        /// <param name="name">name of the localized value</param>
        /// <param name="value">localized value</param>
        public void Add(string name, string value)
        {
            _list.Add(name, value);
        }

        /// <summary>
        /// Remove localized value
        /// </summary>
        /// <param name="name">name of the localized value</param>
        public void Remove(string name)
        {
            _list.Remove(name);
        }

        #endregion

        #region  IEnumerable<KeyValuePair<string,string>>

        /// <summary>
        /// localized values as pair array
        /// </summary>
        /// <returns>enumerator instance</returns>
        public IEnumerator<KeyValuePair<string, string>> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        /// <summary>
        /// localized values as pair array
        /// </summary>
        /// <returns>enumerator instance</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }

    /// <summary>
    /// Represents a collection with language-based values for a NetOffice ToolsDialog instance
    /// </summary>
    public class DialogLocalizationCollection : IEnumerable<DialogLocalization>
    {
        #region Fields

        private List<DialogLocalization> _list;

        #endregion

        #region Ctor
        
        /// <summary>
        /// Creates an instance of the class with default localization(1033, 1031)
        /// </summary>
        /// <param name="name">name of the dialog</param>
        internal DialogLocalizationCollection(string name)
        {
            if (String.IsNullOrEmpty(name))
                throw new ArgumentNullException("name");

            string path1033 = String.Format("{0}.{1}.xml", name, "1033");
            string path1031 = String.Format("{0}.{1}.xml", name, "1031");
            
            XmlDocument document1033 = ReadDefaultLocalization(path1033);
            XmlDocument document1031 = ReadDefaultLocalization(path1031);

            _list = new List<DialogLocalization>();
            _list.Add(CreateDialogLocalization(document1033));
            _list.Add(CreateDialogLocalization(document1031));
           
            Name = name;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Name of the dialog
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// Get a dialog localization subset
        /// </summary>
        /// <param name="lcid">target language id</param>
        /// <returns>DialogLocalization instance or null if lcid not found</returns>
        public DialogLocalization this[int lcid]
        {
            get
            {
                foreach (DialogLocalization item in _list)
                {
                    if (item.LCID == lcid)
                        return item;
                }
                return null;
            }
        }

        /// <summary>
        /// Get a dialog localization subset
        /// </summary>
        /// <param name="lcid">target language id</param>
        /// <param name="firstIfNotFound">return first element(1033) in collection if target is not found, otherwise return null</param>
        /// <returns>DialogLocalization instance or null if lcid not found</returns>
        public DialogLocalization this[int lcid, bool firstIfNotFound]
        {
            get
            {
                foreach (DialogLocalization item in _list)
                {
                    if (item.LCID == lcid)
                        return item;
                }
                return true == firstIfNotFound ? _list[0] : null;
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// Add new language to dialog localization. The new language use 1031(en-us) as content template
        /// </summary>
        /// <param name="lcid">language id for the new language</param>
        /// <returns>new localization instance</returns>
        public DialogLocalization Add(int lcid)
        {
            if (GetLCIDExists(lcid))
                throw new ArgumentException("Duplicate language id");

            IEnumerable<KeyValuePair<string, string>> templates = this[1031];
            DialogLocalization loc = new DialogLocalization(lcid, templates);
            _list.Add(loc);
            return loc;
        }

        /// <summary>
        /// Add new language to dialog localization. The new language use tlcid as content template
        /// </summary>
        /// <param name="lcid">language id for the new language</param>
        /// <param name="tlcid">id of the template language</param>
        /// <returns>new localization instance</returns>
        public DialogLocalization Add(int lcid, int tlcid)
        {
            if (GetLCIDExists(lcid))
                throw new ArgumentException("Duplicate language id");
            if (!GetLCIDExists(tlcid))
                throw new ArgumentException("Template language is missing");

            IEnumerable<KeyValuePair<string, string>> templates = this[tlcid];
            DialogLocalization loc = new DialogLocalization(lcid, templates);
            _list.Add(loc);
            return loc;
        }

        /// <summary>
        /// Remove language from dialog localization
        /// </summary>
        /// <param name="lcid">target language id</param>
        public void Remove(int lcid)
        {
            if (lcid == 1033 || lcid == 1031)
                throw new InvalidOperationException("Unable to remove default language");
            
            DialogLocalization item =  this[lcid];
            if (null == item)
                throw new ArgumentOutOfRangeException("lcid");

            _list.Remove(item);
        }

        /// <summary>
        /// Creates a new DialogLocalization instance with values given in XmlDocument instance
        /// </summary>
        /// <param name="document">content document</param>
        /// <returns>new DialogLocalization instance</returns>
        internal static DialogLocalization CreateDialogLocalization(XmlDocument document)
        {
            if (null == document)
                throw new ArgumentNullException();

            int lcid = Convert.ToInt32(document.FirstChild.Attributes["LCID"].Value);
            Dictionary<string, string> values = new Dictionary<string, string>();
            foreach (XmlNode item in document.FirstChild.FirstChild.ChildNodes)
            {
                string name = item.ChildNodes[0].InnerText;
                string value = item.ChildNodes[1].InnerText;
                values.Add(name, value);
            }

            DialogLocalization result = new DialogLocalization(lcid, values);
            return result;
        }
           
        /// <summary>
        /// Read xml resource localization
        /// </summary>
        /// <param name="resourceAddress">relative resource path</param>
        /// <returns>xml ressource document</returns>
        internal static XmlDocument ReadDefaultLocalization(string resourceAddress)
        {
            Type toolsType = typeof(ToolsDialog);
            string fullQualifiedAddress = String.Format("{0}.{1}", toolsType.Namespace, resourceAddress);
            Stream resourceStream = toolsType.Assembly.GetManifestResourceStream(fullQualifiedAddress);
            if (null == resourceStream)
                throw new IOException("Error accessing resource Stream.");

            XmlDocument document = new XmlDocument();
            document.Load(resourceStream);
            resourceStream.Close();
            resourceStream.Dispose();
            return document;
        }

        private bool GetLCIDExists(int lcid)
        {
            foreach (DialogLocalization item in _list)
            {
                if (item.LCID == lcid)
                    return true;
            }
            return false;
        }

        #endregion

        #region  IEnumerable<DialogLocalization>

        /// <summary>
        /// Returns an enumerator to iterate the instance
        /// </summary>
        /// <returns>enumerator instance</returns>
        public IEnumerator<DialogLocalization> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator to iterate the instance
        /// </summary>
        /// <returns>enumerator instance</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }

    /// <summary>
    /// Contains NetOffice default dialogs localization settings
    /// </summary>
    public class DialogLocalizationSettings : IEnumerable<DialogLocalizationCollection>
    {
        #region Fields

        private List<DialogLocalizationCollection> _list;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="dialogs">dialog definitions as name, resourcePath</param>
        internal DialogLocalizationSettings(IEnumerable<string> dialogs)
        {
            if (null == dialogs)
                throw new ArgumentNullException("dialogs");

            _list = new List<DialogLocalizationCollection>();
            foreach (string item in dialogs)
                Add(item);
        }

        #endregion

        #region Properties

        /// <summary>
        /// Get current dialog localization collection by name
        /// </summary>
        /// <param name="name">name of the dialog</param>
        /// <returns>localization instance</returns>
        public DialogLocalizationCollection this[string name]
        {
            get
            {
                foreach (DialogLocalizationCollection item in _list)
                {
                    if (item.Name == name)
                        return item;
                }
                throw new ArgumentOutOfRangeException("name");
            }
        }

        #endregion

        #region Methods

        private DialogLocalizationCollection Add(string name)
        {
            DialogLocalizationCollection item =  new DialogLocalizationCollection(name);
            _list.Add(item);
            return item;
        }

        private void Remove(DialogLocalizationCollection item)
        {
            _list.Remove(item);
        }

        #endregion

        #region IEnumerable<DialogLocalizationCollection>

        /// <summary>
        /// Returns an enumerator to iterate the instance
        /// </summary>
        /// <returns>enumerator instance</returns>
        public IEnumerator<DialogLocalizationCollection> GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        /// <summary>
        /// Returns an enumerator to iterate the instance
        /// </summary>
        /// <returns>enumerator instance</returns>
        System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator()
        {
            return _list.GetEnumerator();
        }

        #endregion
    }
}
