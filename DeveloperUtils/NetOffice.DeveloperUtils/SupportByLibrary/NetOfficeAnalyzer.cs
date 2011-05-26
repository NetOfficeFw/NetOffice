using System;
using System.Reflection;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using System.Text;
using System.Xml;
using System.Xml.Linq;
using System.Xml.XPath;

using Mono.Cecil;

namespace NetOffice.DeveloperUtils.SupportByLibrary
{
    public static class NetOfficeAnalyzer
    {
        private static List<XDocument> _docuFiles = new List<XDocument>();

        public static string AnalyzeNetOfficeAssemblies(XDocument assemblyDocument, List<AssemblyNameReference> listReferences, AssemblyAnalyzerSettings settings)
        {
            List<string> listResult = new List<string>();

            string path = "";
            if (System.Diagnostics.Debugger.IsAttached)
            {
                path = System.Windows.Forms.Application.StartupPath;
                path = path.Substring(0, path.LastIndexOf("\\"));
                path = path.Substring(0, path.LastIndexOf("\\"));
                path = path.Substring(0, path.LastIndexOf("\\"));
                path = Path.Combine(path, "Docu Files\\");
            }
            else
                path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Docu Files\\");

            foreach (AssemblyNameReference item in listReferences)
            {
                string fileName = Path.Combine(path, item.Name + ".xml");
                XDocument newDocFile = XDocument.Load(fileName);
                _docuFiles.Add(newDocFile);
            }

            foreach (XElement itemClass in assemblyDocument.Element("Assembly").Element("Classes").Elements("Class"))
            {
                foreach (XElement itemField in itemClass.Element("Fields").Elements("Field"))
                {
                    string libName = "";
                    string[] libs = GetSupportByLibrary(itemField, ref libName);
                    if (!FieldPassed(libName, libs, settings))
                    {
                        string warning = string.Format("class {0}: Field {1} {2}; SupportByLibrary {4}, {3}", 
                                                        itemClass.Attribute("Name").Value, itemField.Attribute("Type").Value,
                                                        itemField.Attribute("Name").Value, ToString(libs), libName);
                        listResult.Add(warning); 
                    }
                }

                foreach (XElement itemProperty in itemClass.Element("Properties").Elements("Property"))
                {
                    string libName = "";
                    string[] libs = GetSupportByLibrary(itemProperty, ref libName);
                    if (!FieldPassed(libName, libs, settings))
                    {
                        string warning = string.Format("class {0}: Property {1} {2}; SupportByLibrary {4}, {3}",
                                                        itemClass.Attribute("Name").Value, itemProperty.Attribute("Type").Value,
                                                        itemProperty.Attribute("Name").Value, ToString(libs), libName);
                        listResult.Add(warning);
                    }

                    string[] warnings = new string[0];
                    if (!MethodBodyPassed(itemClass, itemProperty, settings, ref warnings))
                    {
                        foreach (string item in warnings)
                            listResult.Add(item);
                    }

                }

                foreach (XElement itemMethod in itemClass.Element("Methods").Elements("Method"))
                {
                    string[] warnings = new string[0];
                    if (!MethodBodyPassed(itemClass, itemMethod, settings, ref warnings))
                    {

                        foreach (string item in warnings)
                            listResult.Add(item);
                    }
                }
            }

            string result = "";
            foreach (string item in listResult)
                result += item + Environment.NewLine;

            if (0 == listResult.Count)
            { 
                if(0==listReferences.Count)
                    result += "Assembly doesnt use NetOffice." + Environment.NewLine;
                else
                    result += "Assembly works fine with all specified versions." + Environment.NewLine;

            }
                
            return result;
        }

        private static string ToString(string[] value)
        {
            string result = "";
            foreach (string item in value)
                result += item +",";

            result = result.Substring(0, result.Length - 1);
            return result;
        }

        private static XDocument GetDocument(string name)
        {
            foreach (XDocument item in _docuFiles)
            {
                if (name == item.Element("doc").Element("assembly").Value)
                    return item;
            }
            throw (new ArgumentException(name + " not exists."));
        }

        private static string[] GetSupportByLibrary(XElement itemField, ref string libName)
        {
            string[] splitArray = itemField.Attribute("Type").Value.Split(new string[]{ "." },StringSplitOptions.RemoveEmptyEntries);
            XDocument apiDocument = GetDocument(splitArray[1]);

            XElement memberNode = (from a in apiDocument.Element("doc").Element("members").Elements("member")
                                where a.Attribute("name").Value.Equals("T:" + itemField.Attribute("Type").Value)
                                select a).FirstOrDefault();
            
            string value = memberNode.Value.Replace("\r\n", "").Replace("\n", "").Replace("\t", "");
            value = value.Substring(value.IndexOf("SupportByLibrary") + "SupportByLibrary".Length).Trim();
            libName = value.Substring(0, value.IndexOf(" ")).Replace(",","");
            value = value.Substring(value.IndexOf(" ") + 1).Replace(" ", "");

            string[] returnArray = value.Split(new String[]{","},StringSplitOptions.RemoveEmptyEntries);
            return returnArray;
        }

        private static bool Includes(string[] libs, string value)
        {
            foreach (string item in libs)
            {
                if (item == value)
                    return true;
            }
            return false;
        }

        private static bool CheckLibAttribute(string[] libs, AssemblyAnalyzerSettingsLibrary libSettings)
        {
            if( (libSettings.Version9) &&  (!Includes(libs, "9")) ) 
                    return false;
            
            if ((libSettings.Version10) && (!Includes(libs, "10")))
                return false;

            if ((libSettings.Version11) && (!Includes(libs, "11")))
                return false;
            
            if ((libSettings.Version12) && (!Includes(libs, "12")))
                return false;

            if ((libSettings.Version14) && (!Includes(libs, "14")))
                return false;

            return true;
        }

        private static bool FieldPassed(string name, string[] libs, AssemblyAnalyzerSettings settings)
        {
            switch (name)
            {
                case "Excel":
                    if (!CheckLibAttribute(libs, settings.Excel))
                        return false;
                    break;
                case "Word":
                    if (!CheckLibAttribute(libs, settings.Word))
                        return false;
                    break;
                case "Outlook":
                    if (!CheckLibAttribute(libs, settings.Outlook))
                        return false;
                    break;
                case "PowerPoint":
                    if (!CheckLibAttribute(libs, settings.PowerPoint))
                        return false;
                    break;
                case "Access":
                    if (!CheckLibAttribute(libs, settings.Access))
                        return false;
                    break;
                case "Office":
                    if (!CheckLibAttribute(libs, settings.Office))
                        return false;
                    break;
            }

            return true;
        }

        private static bool IsEnum(string type)
        {
            string[] splitArray = type.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            if (splitArray[splitArray.Length - 2] == "Enums")
                return true;
            else
                return false;
        }

        private static bool MethodBodyPassed(XElement itemClass, XElement itemMethod, AssemblyAnalyzerSettings settings, ref string[] warnings)
        {
            List<string> listWarnings = new List<string>();

            foreach (XElement itemEntity in itemMethod.Elements())
            {
                switch (itemEntity.Name.LocalName)
                {
                    case "ReturnValue":
                    {
                        string libName = "";
                        string[] libs = GetSupportByLibrary(itemEntity, ref libName);
                        if (!FieldPassed(libName, libs, settings))
                        {
                            string warning = string.Format("class {0}: Entity ReturnValue {4} {1}; SupportByLibrary {3}, {2}",
                                                            itemClass.Attribute("Name").Value,
                                                            itemMethod.Attribute("Name").Value,
                                                            ToString(libs), libName, itemEntity.Attribute("Type").Value);
                            
                            listWarnings.Add(warning);                          
                        }
                        break;
                    }
                    case "Var":
                    {
                        string libName = "";
                        string[] libs = GetSupportByLibrary(itemEntity, ref libName);
                        if (!FieldPassed(libName, libs, settings))
                        {
                            string warning = string.Format("class {0}: Entity Local Variable {4} {1}; SupportByLibrary {3}, {2}",
                                                            itemClass.Attribute("Name").Value,
                                                            itemMethod.Attribute("Name").Value,
                                                            ToString(libs), libName, itemEntity.Attribute("Type").Value);

                            listWarnings.Add(warning);
                        }

                        string type = itemEntity.Attribute("Type").Value;
                        foreach (XElement itemSet in itemEntity.Elements("Set"))
                        {
                            string setValue = itemSet.Attribute("Value").Value;
                            if (IsEnum(type))
                            { 
                            }
                        }

                        break;
                    }
                    case "FieldSet":
                        break;

                    case "Call":
                        break;

                }
            }

            return false;
        }
    }
}
