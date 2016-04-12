using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using HtmlAgilityPack;

namespace NOBuildTools.ReferenceAnalyzer
{
    /// <summary>
    /// Progress log action handler
    /// </summary>
    /// <param name="message">log action message</param>
    public delegate void LogAction(string message);

    internal static class Parser
    {
        #region Fields

        private static string _rootAdress = "http://msdn.microsoft.com/en-us";
        
        private static string _excelTypesRelative = "/library/office/ff194068.aspx";
        private static string _excelEnumsRelative = "/library/office/ff838815.aspx";

        private static string _accessTypesRelative = "/library/office/ff192120.aspx";
        private static string _accessEnumsRelative = "/library/office/jj713155.aspx";
        private static string _accessConstantsRelative = "/library/office/jj713057.aspx";

        private static string _officeTypesRelative = "/library/office/ff861484.aspx";
        private static string _officeEnumsRelative = "/library/office/jj229676.aspx";

        private static string _outlookTypesRelative = "/library/office/ff866465.aspx";
        private static string _outlookEnumsRelative = "/library/office/ff860961.aspx";

        private static string _powerPointTypesRelative = "/library/office/ff743835.aspx";
        private static string _powerPointEnumsRelative = "/library/office/ff744042.aspx";

        private static string _projectTypesRelative = "/library/office/ff920539(v=office.14).aspx";
        private static string _projectEnumsRelative = "/library/office/ff920788(v=office.14).aspx";

        private static string _visioTypesRelative = "/library/ff765377(v=office.14).aspx";
        private static string _visioEnumsRelative = "/library/ff769457(v=office.14).aspx";

        private static string _wordTypesRelative = "/library/office/ff837519.aspx";
        private static string _wordEnumsRelative = "/library/office/dn353221.aspx";

        #endregion

        #region Parse Word

        /// <summary>
        /// Parse Word Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseWord(XDocument document, LogAction func)
        {
            XElement WordNode = new XElement("Word");
            (document.FirstNode as XElement).Add(WordNode);
            ParseWordTypes(WordNode, func);
            ParseWordEnums(WordNode, func);
            ParseWordTypesMembers(WordNode, func);
        }

        private static void ParseWordTypes(XElement excelNode, LogAction func)
        {
            func("Parse Word Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _wordTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Word Types recieved", counter));
        }

        private static void ParseWordEnums(XElement excelNode, LogAction func)
        {
            func("Parse Word Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _wordEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Word Enums recieved", counter));
        }

        private static void ParseWordTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Word Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParseWordTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseWordTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseWordTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseWordTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseWordTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseWordTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseWordTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Visio

        /// <summary>
        /// Parse Visio Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseVisio(XDocument document, LogAction func)
        {
            XElement VisioNode = new XElement("Visio");
            (document.FirstNode as XElement).Add(VisioNode);
            ParseVisioTypes(VisioNode, func);
            ParseVisioEnums(VisioNode, func);
            ParseVisioTypesMembers(VisioNode, func);
        }

        private static void ParseVisioTypes(XElement excelNode, LogAction func)
        {
            func("Parse Visio Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _visioTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Visio Types recieved", counter));
        }

        private static void ParseVisioEnums(XElement excelNode, LogAction func)
        {
            func("Parse Visio Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _visioEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Visio Enums recieved", counter));
        }

        private static void ParseVisioTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Visio Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParseVisioTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseVisioTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseVisioTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseVisioTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseVisioTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseVisioTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseVisioTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Project

        /// <summary>
        /// Parse MSProject Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseProject(XDocument document, LogAction func)
        {
            XElement projectNode = new XElement("MSProject");
            (document.FirstNode as XElement).Add(projectNode);
            ParseProjectTypes(projectNode, func);
            ParseProjectEnums(projectNode, func);
            ParseProjectTypesMembers(projectNode, func);
        }

        private static void ParseProjectTypes(XElement excelNode, LogAction func)
        {
            func("Parse Project Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _projectTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Project Types recieved", counter));
        }

        private static void ParseProjectEnums(XElement excelNode, LogAction func)
        {
            func("Parse Project Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _projectEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Project Enums recieved", counter));
        }

        private static void ParseProjectTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Project Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParseProjectTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseProjectTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseProjectTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseProjectTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseProjectTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseProjectTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseProjectTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse PowerPoint

        /// <summary>
        /// Parse PPoint Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParsePowerPoint(XDocument document, LogAction func)
        {
            XElement pPointNode = new XElement("PowerPoint");
            (document.FirstNode as XElement).Add(pPointNode);
            ParsePowerPointTypes(pPointNode, func);
            ParsePowerPointEnums(pPointNode, func);
            ParsePowerPointTypesMembers(pPointNode, func);
        }

        private static void ParsePowerPointTypes(XElement excelNode, LogAction func)
        {
            func("Parse PowerPoint Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _powerPointTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} PowerPoint Types recieved", counter));
        }

        private static void ParsePowerPointEnums(XElement excelNode, LogAction func)
        {
            func("Parse PowerPoint Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _powerPointEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} PowerPoint Enums recieved", counter));
        }

        private static void ParsePowerPointTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse PowerPoint Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParsePowerPointTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParsePowerPointTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParsePowerPointTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParsePowerPointTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParsePowerPointTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParsePowerPointTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParsePowerPointTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Outlook

        /// <summary>
        /// Parse Outlook Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseOutlook(XDocument document, LogAction func)
        {
            XElement outlookNode = new XElement("Outlook");
            (document.FirstNode as XElement).Add(outlookNode);
            ParseOutlookTypes(outlookNode, func);
            ParseOutlookEnums(outlookNode, func);
            ParseOutlookTypesMembers(outlookNode, func);
        }

        private static void ParseOutlookTypes(XElement excelNode, LogAction func)
        {
            func("Parse Outlook Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _outlookTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Outlook Types recieved", counter));
        }

        private static void ParseOutlookEnums(XElement excelNode, LogAction func)
        {
            func("Parse Outlook Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _outlookEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Outlook Enums recieved", counter));
        }

        private static void ParseOutlookTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Outlook Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParseOutlookTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOutlookTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOutlookTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOutlookTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseOutlookTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseOutlookTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseOutlookTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Office

        /// <summary>
        /// Parse Common Office Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseOffice(XDocument document, LogAction func)
        {
            XElement officeNode = new XElement("Office");
            (document.FirstNode as XElement).Add(officeNode);
            ParseOfficeTypes(officeNode, func);
            ParseOfficeEnums(officeNode, func);
            ParseOfficeTypesMembers(officeNode, func);
        }

        private static void ParseOfficeTypes(XElement excelNode, LogAction func)
        {
            func("Parse Office Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _officeTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Office Types recieved", counter));
        }

        private static void ParseOfficeEnums(XElement excelNode, LogAction func)
        {
            func("Parse Office Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _officeEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Office Enums recieved", counter));
        }

        private static void ParseOfficeTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Office Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseOfficeTypeMembers(item, func);
            }
        }

        private static void ParseOfficeTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOfficeTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOfficeTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseOfficeTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseOfficeTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseOfficeTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseOfficeTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Excel

        /// <summary>
        /// Parse Excel Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseExcel(XDocument document, LogAction func)
        {
            XElement excelNode = new XElement("Excel");
            (document.FirstNode as XElement).Add(excelNode);
            ParseExcelTypes(excelNode, func);
            ParseExcelEnums(excelNode, func);
            ParseExcelTypesMembers(excelNode, func);
        }

        private static void ParseExcelTypes(XElement excelNode, LogAction func)
        {
            func("Parse Excel Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _excelTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                 string pageContent = DownloadPage(client, excelRootReferencePage);
                 HtmlDocument doc = new HtmlDocument();
                 doc.LoadHtml(pageContent);
                 var root = doc.DocumentNode;

                 var divNodes = root.Descendants("div").ToList();
                 foreach (var item in divNodes)
                 {
                     string className = item.GetAttributeValue("class", null);
                     if (className == "toclevel2")
                     {
                         string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                         string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                         if (null != href && null != name)
                         {
                             if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                             {
                                 name = name.Substring(0, name.Length - " Object".Length);
                                 rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                 counter++;
                             }
                         }
                     }

                 }
            }

            func(String.Format("{0} Excel Types recieved", counter));
        }
    
        private static void ParseExcelEnums(XElement excelNode, LogAction func)
        {
            func("Parse Excel Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);
           
            int counter = 0;
            string excelRootReferencePage = _rootAdress + _excelEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);
                            rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            counter++;
                        }
                    }

                }

            }

            func(String.Format("{0} Excel Enums recieved", counter));
        }
        
        private static void ParseExcelTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Excel Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseExcelTypeMembers(item, func);
            }
        }

        private static void ParseExcelTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

             using (var client = new System.Net.WebClient())
             {
                 string pageLink = typeNode.Element("Link").Value;
                 string pageContent = DownloadPage(client, pageLink);
                 HtmlDocument doc = new HtmlDocument();
                 doc.LoadHtml(pageContent);
                 var root = doc.DocumentNode;

                 var divNodes = root.Descendants("div").ToList();
                 foreach (var item in divNodes)
                 {
                      string className = item.GetAttributeValue("class", null);
                      if (className == "toclevel2")
                      {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseExcelTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseExcelTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseExcelTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                      }
                 }
             }
        }

        private static void ParseExcelTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                     string className = item.GetAttributeValue("class", null);
                     if (className == "toclevel2")
                     {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                     }
                }
            }
        }
        
        private static void ParseExcelTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseExcelTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        #endregion

        #region Parse Access 

        /// <summary>
        /// Parse Access Docu pages
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static void ParseAccess(XDocument document, LogAction func)
        {
            XElement accessNode = new XElement("Access");
            (document.FirstNode as XElement).Add(accessNode);
            ParseAccessTypes(accessNode, func);
            ParseAccessEnums(accessNode, func);
            ParseAccessConstants(accessNode, func);
            ParseAccessTypesMembers(accessNode, func);
        }

        private static void ParseAccessTypes(XElement excelNode, LogAction func)
        {
            func("Parse Access Types");

            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _accessTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.EndsWith(" Object", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Object".Length);
                                rootNode.Add(new XElement("Type", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }
            }

            func(String.Format("{0} Access Types recieved", counter));
        }

        private static void ParseAccessEnums(XElement excelNode, LogAction func)
        {
            func("Parse Access Enums");

            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            int counter = 0;
            string excelRootReferencePage = _rootAdress + _accessEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = DownloadPage(client, excelRootReferencePage);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.Substring(0, name.Length - " Enumeration".Length);

                            if (!name.Equals("OldConstants", StringComparison.InvariantCultureIgnoreCase))
                            { 
                                rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                                counter++;
                            }
                        }
                    }

                }

            }

            func(String.Format("{0} Access Enums recieved", counter));
        }

        private static void ParseAccessTypesMembers(XElement typeNode, LogAction func)
        {
            func("Parse Access Type Members");
            foreach (XElement item in typeNode.Element("Types").Elements("Type"))
            {
                ParseAccessTypeMembers(item, func);
            }
        }

        private static void ParseAccessTypeMembers(XElement typeNode, LogAction func)
        {
            XElement propsNode = new XElement("Properties");
            XElement methodsNode = new XElement("Methods");
            XElement eventsNode = new XElement("Events");
            typeNode.Add(propsNode);
            typeNode.Add(methodsNode);
            typeNode.Add(eventsNode);

            using (var client = new System.Net.WebClient())
            {
                string pageLink = typeNode.Element("Link").Value;
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;

                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            name = name.ToLower().Trim();
                            switch (name)
                            {
                                case "properties":
                                    propsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseAccessTypeProperties(propsNode, func);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseAccessTypeMethods(methodsNode, func);
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    ParseAccessTypeEvents(eventsNode, func);
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }
        }

        private static void ParseAccessTypeProperties(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Property", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseAccessTypeMethods(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Method", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }

        private static void ParseAccessTypeEvents(XElement propertiesNode, LogAction func)
        {
            using (var client = new System.Net.WebClient())
            {
                string pageLink = XmlConvert.DecodeName(propertiesNode.Attribute("Link").Value);
                string pageContent = DownloadPage(client, pageLink);
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(pageContent);
                var root = doc.DocumentNode;
                var divNodes = root.Descendants("div").ToList();
                foreach (var item in divNodes)
                {
                    string className = item.GetAttributeValue("class", null);
                    if (className == "toclevel2")
                    {
                        string href = item.FirstChild.NextSibling.GetAttributeValue("href", null);
                        string name = item.FirstChild.NextSibling.GetAttributeValue("title", null);
                        if (null != href && null != name)
                        {
                            if (name.IndexOf(" ") > -1)
                                name = name.Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries)[0];
                            propertiesNode.Add(new XElement("Event", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            func("");
                        }
                    }
                }
            }
        }
        
        private static void ParseAccessConstants(XElement excelNode, LogAction func)
        {
            func("Parse Access Constants");
            XElement constantNode = new XElement("OldConstants", _rootAdress + _accessConstantsRelative);
            excelNode.Add(constantNode);
            func(String.Format("{0} Access Contants recieved", 1));
        }

        #endregion

        #region Methods

        /// <summary>
        /// Parse MSDN Docu pages for MS-Office
        /// </summary>
        /// <param name="document">document to fill</param>
        /// <param name="func">progress handler</param>
        internal static XDocument ParseReference(LogAction func)
        {
            func("Parse References ");
            XDocument document = new XDocument();
            document.Add(new XElement("NOBuildTools.ReferenceAnalyzer"));
            ParseExcel(document, func);
            ParseAccess(document, func);
            ParseOffice(document, func);
            ParseOutlook(document, func);
            ParsePowerPoint(document, func);
            ParseProject(document, func);
            ParseVisio(document, func);
            ParseWord(document, func);
            func("Done!");

            return document;
        }

        private static string DownloadPage(System.Net.WebClient client, string uri)
        {
            try
            {
                string pageContent = client.DownloadString(uri);
                return pageContent;
            }
            catch
            {
                return DownloadPage(client, uri);
            }
        }

        #endregion

    }
}
