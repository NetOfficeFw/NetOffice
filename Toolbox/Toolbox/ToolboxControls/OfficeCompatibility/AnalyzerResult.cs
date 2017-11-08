using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Mono.Cecil;
using Mono.Cecil.Cil;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeCompatibility
{
    /// <summary>
    /// Generaly info about assembly usage to an MS-Office product in specific version
    /// </summary>
    public enum SupportVersion
    {
        /// <summary>
        /// assembly doesnt the ms office product version
        /// </summary>
        NotUse = 0,

        /// <summary>
        /// assembly use the ms office product and requested features is full supported in all versions 
        /// </summary>
        Support = 1,

        /// <summary>
        /// assembly use the ms office product and requested features only partialy supported in one or few versions
        /// </summary>
        NotSupport = 2
    }

    /// <summary>
    /// Generaly info about assembly usage to an MS-Office product in specific version
    /// </summary>
    public class SupportInfo
    {
        #region Fields

        private SupportVersion _support;
        private int _version;
        private string _name;

        #endregion

        #region Construction

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="support">generaly version support</param>
        /// <param name="name">name of the office product</param>
        /// <param name="version">office product version number</param>
        internal SupportInfo(SupportVersion support, string name, int version)
        {
            _support = support;
            _name = name;
            _version = version;
        }

        #endregion

        #region Properties

        /// <summary>
        /// Version support
        /// </summary>
        public SupportVersion Support
        {
            get
            {
                return _support;
            }
        }

        /// <summary>
        /// Office product version number
        /// </summary>
        public int Version
        {
            get
            {
                return _version;
            }
        }

        /// <summary>
        /// Name of the office product
        /// </summary>
        public string Name
        {
            get
            {
                return _name;
            }
        }

        #endregion
    }

    public class AnalyzerResult
    {
        #region Fields

        private bool _containsNetOfficeReferences;
        private XDocument _report;
        private SupportInfo[] _office;
        private SupportInfo[] _excel;
        private SupportInfo[] _word;
        private SupportInfo[] _outlook;
        private SupportInfo[] _powerPoint;
        private SupportInfo[] _access;
        private SupportInfo[] _project;
        private SupportInfo[] _visio;
        private SupportInfo[] _publisher;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="containsNetOfficeReferences">analyzed assembly contains NetOffice requests</param>
        internal AnalyzerResult(bool containsNetOfficeReferences)
        {
            _containsNetOfficeReferences = containsNetOfficeReferences;
        }

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        /// <param name="report">detailed usage report</param>
        internal AnalyzerResult(XDocument report)
        {
            _report = report;
            _containsNetOfficeReferences = true;
            _office = new SupportInfo[7];
            _excel = new SupportInfo[7];
            _word = new SupportInfo[7];
            _outlook = new SupportInfo[7];
            _powerPoint = new SupportInfo[7];
            _access = new SupportInfo[7];
            _project = new SupportInfo[7];
            _visio = new SupportInfo[7];
            _publisher = new SupportInfo[7];

            RemoveDelegateTypes();

            SetupSupportInfo(_office, "Office");
            SetupSupportInfo(_excel, "Excel");
            SetupSupportInfo(_word, "Word");
            SetupSupportInfo(_outlook, "Outlook");
            SetupSupportInfo(_powerPoint, "PowerPoint");
            SetupSupportInfo(_access, "Access");
            SetupSupportInfo(_project, "MSProject");
            SetupSupportInfo(_visio, "Visio");
            SetupSupportInfo(_publisher, "Publisher");
        }

        #endregion

        #region Properties

        /// <summary>
        /// Detailed usage report
        /// </summary>
        public XDocument Report
        {
            get
            {
                return _report;
            }
        }

        /// <summary>
        /// Analyzed assembly contains NetOffice requests
        /// </summary>
        public bool ContainsNetOfficeReferences
        {
            get
            {
                return _containsNetOfficeReferences;
            }
        }

        /// <summary>
        /// Common Office support info
        /// </summary>
        public SupportInfo[] Office
        {
            get
            {
                return _office;
            }
        }

        /// <summary>
        /// Excel support info
        /// </summary>
        public SupportInfo[] Excel
        {
            get
            {
                return _excel;
            }
        }

        /// <summary>
        /// Word support info
        /// </summary>
        public SupportInfo[] Word
        {
            get
            {
                return _word;
            }
        }

        /// <summary>
        /// Outlook support info
        /// </summary>
        public SupportInfo[] Outlook
        {
            get
            {
                return _outlook;
            }
        }

        /// <summary>
        /// PowerPoint support info
        /// </summary>
        public SupportInfo[] PowerPoint
        {
            get
            {
                return _powerPoint;
            }
        }

        /// <summary>
        /// Access support info
        /// </summary>
        public SupportInfo[] Access
        {
            get
            {
                return _access;
            }
        }

        /// <summary>
        /// Project support info 
        /// </summary>
        public SupportInfo[] Project
        {
            get
            {
                return _project;
            }
        }

        /// <summary>
        /// Visio support info
        /// </summary>
        public SupportInfo[] Visio
        {
            get
            {
                return _visio;
            }
        }

        #endregion

        #region Methods

        private void SetupSupportInfo(SupportInfo[] info, string name)
        {
            bool inUse = false;

            bool has09Support = true;
            bool has10Support = true;
            bool has11Support = true;
            bool has12Support = true;
            bool has14Support = true;
            bool has15Support = true;
            bool has16Support = true;

            foreach (XElement item in _report.Element("Document").Element("Assembly").Element("Classes").Elements("Class"))
            {

                var supportNodes = (from a in item.Descendants("SupportByLibrary")
                                    where a.Attribute("Api").Value.Equals(name, StringComparison.InvariantCultureIgnoreCase)
                                    select a);

                if ((null == supportNodes) || (supportNodes.Count() == 0))
                    continue;

                inUse = true;

                foreach (XElement typeNodeItem in supportNodes)
                {
                    bool find09Support = IncludesVersion(typeNodeItem, "9");
                    bool find10Support = IncludesVersion(typeNodeItem, "10");
                    bool find11Support = IncludesVersion(typeNodeItem, "11");
                    bool find12Support = IncludesVersion(typeNodeItem, "12");
                    bool find14Support = IncludesVersion(typeNodeItem, "14");
                    bool find15Support = IncludesVersion(typeNodeItem, "15");
                    bool find16Support = IncludesVersion(typeNodeItem, "16");

                    if (!find09Support)
                        has09Support = false;
                    if (!find10Support)
                        has10Support = false;
                    if (!find11Support)
                        has11Support = false;
                    if (!find12Support)
                        has12Support = false;
                    if (!find14Support)
                        has14Support = false;
                    if (!find15Support)
                        has15Support = false;
                    if (!find16Support)
                        has16Support = false;
                }
            }

            if (inUse)
            {
                info[0] = new SupportInfo(BoolToSupportVersion(has09Support), name, 9);
                info[1] = new SupportInfo(BoolToSupportVersion(has10Support), name, 10);
                info[2] = new SupportInfo(BoolToSupportVersion(has11Support), name, 11);
                info[3] = new SupportInfo(BoolToSupportVersion(has12Support), name, 12);
                info[4] = new SupportInfo(BoolToSupportVersion(has14Support), name, 14);
                info[5] = new SupportInfo(BoolToSupportVersion(has15Support), name, 15);
                info[6] = new SupportInfo(BoolToSupportVersion(has16Support), name, 16);
            }
            else
            {
                info[0] = new SupportInfo(SupportVersion.NotUse, name, 9);
                info[1] = new SupportInfo(SupportVersion.NotUse, name, 10);
                info[2] = new SupportInfo(SupportVersion.NotUse, name, 11);
                info[3] = new SupportInfo(SupportVersion.NotUse, name, 12);
                info[4] = new SupportInfo(SupportVersion.NotUse, name, 14);
                info[5] = new SupportInfo(SupportVersion.NotUse, name, 15);
                info[6] = new SupportInfo(SupportVersion.NotUse, name, 16);
            }
        }

        private void RemoveDelegateTypes()
        {
            List<XElement> listToDelete = new List<XElement>();
            var typeNodes = (from a in _report.Descendants("Entity")
                             select a);

            foreach (XElement item in typeNodes)
            {
                if (0 == item.Element("SupportByLibrary").Elements("Version").Count())
                    listToDelete.Add(item);
            }

            foreach (XElement item in listToDelete)
                item.Remove();
        }


        private static bool IncludesVersion(XElement supportByLibraryNode, string version)
        {
            foreach (XElement versionItem in supportByLibraryNode.Elements("Version"))
            {
                string versionSupport = versionItem.Value;
                if (versionSupport == version)
                    return true;
            }

            return false;
        }

        private static SupportVersion BoolToSupportVersion(bool value)
        {
            if (true == value)
                return SupportVersion.Support;
            else
                return SupportVersion.NotSupport;
        }

        #endregion
    }
}
