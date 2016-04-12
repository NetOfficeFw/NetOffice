using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Text;

namespace NOBuildTools.VersionUpdater
{  
    /// <summary>
    /// Little helper to read and save config
    /// </summary>
    internal static class ConfigManager
    {
        /// <summary>
        /// Save a configuration file
        /// </summary>
        /// <param name="fullFileName">target config file</param>
        /// <param name="directory">target update root directory</param>
        /// <param name="changeMarker">change Marker</param>
        /// <param name="fromNet">from .net version</param>
        /// <param name="toNet">to .net versions</param>
        /// <param name="changeKeyFiles">change also key files</param>
        /// <param name="keyFilesFolder">root folder with key files</param>
        public static void SaveConfigurationToXMLFile(string fullFileName, string directory, bool changeMarker, string fromNet, string toNet, bool changeKeyFiles, string keyFilesFolder)
        {
            if (File.Exists(fullFileName))
                File.Delete(fullFileName);
            XDocument document = new XDocument(new XElement("NOBuildTools.VersionUpdater"));
            XElement root = document.FirstNode as XElement;

            root.Add(new XElement("TargetFolder", XmlConvert.EncodeName(directory)));
            root.Add(new XElement("ChangeMarker", XmlConvert.EncodeName(changeMarker.ToString())));
            root.Add(new XElement("From", XmlConvert.EncodeName(fromNet)));
            root.Add(new XElement("To", XmlConvert.EncodeName(fromNet)));
            root.Add(new XElement("ChangeKeyFiles", XmlConvert.EncodeName(changeKeyFiles.ToString())));
            root.Add(new XElement("KeyFilesFolder", XmlConvert.EncodeName(keyFilesFolder)));

            document.Save(fullFileName);
        }

        /// <summary>
        /// Loads a configuration file
        /// </summary>
        /// <param name="fullFileName">target config file</param>
        /// <param name="directory">target update root directory</param>
        /// <param name="changeMarker">change Marker</param>
        /// <param name="fromNet">from .net version</param>
        /// <param name="toNet">to .net versions</param>
        /// <param name="changeKeyFiles">change also key files</param>
        /// <param name="keyFilesFolder">root folder with key files</param>
        public static void LoadConfigurationFromConfigFile(string fullFileName, ref string directory, ref bool changeMarker, ref string fromNet, ref string toNet, ref bool changeKeyFiles, ref string keyFilesFolder)
        {
            if (!File.Exists(fullFileName))
                throw new FileNotFoundException(fullFileName);

            XDocument document = XDocument.Load(fullFileName);
            XElement root = document.FirstNode as XElement;
            if (root.Name != "NOBuildTools.VersionUpdater")
                throw new FormatException("Wrong Magic");

            directory = XmlConvert.DecodeName(root.Element("TargetFolder").Value);
            changeMarker = Convert.ToBoolean(XmlConvert.DecodeName(root.Element("ChangeMarker").Value));
            fromNet = XmlConvert.DecodeName(root.Element("From").Value);
            toNet = XmlConvert.DecodeName(root.Element("To").Value);
            changeKeyFiles = Convert.ToBoolean(XmlConvert.DecodeName(root.Element("ChangeKeyFiles").Value));
            keyFilesFolder = XmlConvert.DecodeName(root.Element("KeyFilesFolder").Value);
        }
    }
}
