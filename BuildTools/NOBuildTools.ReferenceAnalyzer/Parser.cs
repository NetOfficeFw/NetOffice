using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;
using System.Text;
using HtmlAgilityPack;

namespace NOBuildTools.ReferenceAnalyzer
{
    internal static class Parser
    {
        private static string _rootAdress = "http://msdn.microsoft.com/en-us";
        private static string _excelTypesRelative = "/library/office/ff194068.aspx";
        private static string _excelEnumsRelative = "/library/office/ff838815.aspx";


        internal static XDocument ParseReference()
        {
            XDocument document = new XDocument();
            document.Add(new XElement("NOBuildTools.ReferenceAnalyzer"));
            ParseExcel(document);
            return document;
        }

        internal static void ParseExcel(XDocument document)
        {
            XElement excelNode = new XElement("Excel");
            (document.FirstNode as XElement).Add(excelNode);
            ParseExcelTypes(document.FirstNode as XElement);
            ParseExcelEnums(document.FirstNode as XElement);
            ParseExcelConstans(document.FirstNode as XElement);
        }

        private static void ParseExcelTypes(XElement excelNode)
        {
            XElement rootNode = new XElement("Types");
            excelNode.Add(rootNode);

            string excelRootReferencePage = _rootAdress + _excelTypesRelative;
            using (var client = new System.Net.WebClient())
            {
                 string pageContent = client.DownloadString(excelRootReferencePage);
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
                             }
                         }
                     }

                 }
            }
        }

        private static void ParseExcelTypeMembers(XElement typeNode)
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
                 string pageContent = client.DownloadString(pageLink);
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
                                    ParseExcelTypeProperties(propsNode);
                                    break;
                                case "methods":
                                    methodsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    break;
                                case "events":
                                    eventsNode.Add(new XAttribute("Link", XmlConvert.EncodeName(_rootAdress + href)));
                                    break;
                                default:
                                    break;
                            }
                        }
                      }
                 }
             }
        }

        private static void ParseExcelTypeProperties(XElement propertiesNode)
        { 
        }

        private static void ParseExcelEnums(XElement excelNode)
        {
            XElement rootNode = new XElement("Enums");
            excelNode.Add(rootNode);

            string excelRootReferencePage = _rootAdress + _excelEnumsRelative;
            using (var client = new System.Net.WebClient())
            {
                string pageContent = client.DownloadString(excelRootReferencePage);
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
                            if (!name.Equals("Constants Enumeration", StringComparison.InvariantCultureIgnoreCase))
                            {
                                name = name.Substring(0, name.Length - " Enumeration".Length);
                                rootNode.Add(new XElement("Enum", new XElement("Name", name), new XElement("Link", _rootAdress + href)));
                            }
                        }
                    }

                }

            }
        }

        private static void ParseExcelConstans(XElement excelNode)
        {
            XElement constantNode = new XElement("Constants", "http://msdn.microsoft.com/en-us/library/office/ff197824.aspx");
            excelNode.Add(constantNode);
        }
    }
}
