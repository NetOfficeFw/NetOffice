using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Text;

namespace NOBuildTools.VersionUpdater
{
    internal static class ConfigManager
    {
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
