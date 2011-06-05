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

        #region Public Methods

       
        public static string AnalyzeNetOfficeAssemblies(XDocument assemblyDocument, List<AssemblyNameReference> listReferences, AssemblyAnalyzerSettings settings)
        {
            List<string> listResult = new List<string>();

            // set docufiles path
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

            // load docu files
            foreach (AssemblyNameReference item in listReferences)
            {
                string fileName = Path.Combine(path, item.Name + ".Description.xml");
                XDocument newDocFile = XDocument.Load(fileName);
                _docuFiles.Add(newDocFile);
            }

            // fetch all classes
            foreach (XElement itemClass in assemblyDocument.Element("Assembly").Element("Classes").Elements("Class"))
            {
                int listPosition = listResult.Count;
                bool notPassed = false;

                // fields, scan field types for library support
                foreach (XElement itemField in itemClass.Element("Fields").Elements("Field"))
                {
                    string libName = "";
                    string[] libs = GetSupportByLibrary(itemField, ref libName);
                    if (!FieldPassed(libName, libs, settings))
                    {
                        notPassed = true;

                        string warning = string.Format("class {0}: {1} {2}; SupportByLibrary {3}", 
                                                        itemClass.Attribute("Name").Value, itemField.Attribute("Type").Value,
                                                        itemField.Attribute("Name").Value, ToString(libs));
                        listResult.Add(warning); 
                    }
                }
                
                // properties
                foreach (XElement itemProperty in itemClass.Element("Properties").Elements("Property"))
                {
                    string libName = "";
                    string[] libs = GetSupportByLibrary(itemProperty, ref libName);
                    if (!FieldPassed(libName, libs, settings))
                    {
                        notPassed = true;
 
                        string warning = string.Format("class {0}: {1} {2}; SupportByLibrary {3}",
                                                                            itemClass.Attribute("Name").Value, itemProperty.Attribute("Type").Value,
                                                                            itemProperty.Attribute("Name").Value, ToString(libs));
                        listResult.Add(warning);
                    }
                    
                    string[] warnings = new string[0];
                    if (!MethodBodyPassed(itemClass, itemProperty, settings, ref warnings, true))
                    {
                        notPassed = true;
 
                        foreach (string item in warnings)
                            listResult.Add(item);
                    }
                     
                }

                // methods
                foreach (XElement itemMethod in itemClass.Element("Methods").Elements("Method"))
                {
                    string[] warnings = new string[0];
                    if (!MethodBodyPassed(itemClass, itemMethod, settings, ref warnings,false))
                    {
                        notPassed = true;

                        foreach (string item in warnings)
                            listResult.Add(item);
                    }
                }

                if (notPassed)
                {
                    string message = "Class " + itemClass.Attribute("Name").Value;
                    message += "\r\n" + Space("=", message.Length);
                    listResult.Insert(listPosition, message);
                }
            }

            string result = "";
            
            if (0 < listResult.Count)
            {
                foreach (string item in listResult)
                    result += item + Environment.NewLine;
            }
            else
            {
                if (0 == listReferences.Count)
                    result += "Assembly doesnt use NetOffice." + Environment.NewLine;
                else
                    result += "Assembly works fine with all specified versions." + Environment.NewLine;
            }

            return result;
        }

        #endregion

        #region Private Methods
       
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
                if (name == item.Element("Assembly").Attribute("Name").Value)
                    return item;
            }
            throw (new ArgumentException(name + " not exists."));
        }

        private static string GetEnumMemberName(string enumName, string value)
        {
            string[] splitArray = enumName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            XDocument apiDocument = GetDocument(splitArray[1]);
            XElement enumNode = (from a in apiDocument.Element("Assembly").Element("Enums").Elements("Enum")
                            where a.Attribute("Name").Value.Equals(splitArray[splitArray.Length-1])
                            select a).FirstOrDefault();


            foreach (XElement item in enumNode.Elements("Member"))
            {
                if (value == item.Attribute("Value").Value)
                    return item.Attribute("Name").Value;
            }
            throw new ArgumentException(enumName + " not exists or has no remarks tag");
        }
        
        private static string[] GetSupportByLibrarySet(XElement itemField, string libName)
        {
            switch (itemField.Name.LocalName)
            {
                case "Set":
                {
                    XDocument apiDocument = GetDocument(libName+"Api");
                    XElement memberNode = (from a in apiDocument.Element("Assembly").Element("Enums").Elements("Enum").Elements("Member")
                                           where a.Attribute("Name").Value.Equals(itemField.Attribute("Value").Value)
                                               select a).FirstOrDefault();
 

                    if(null == memberNode)
                        memberNode = (from a in apiDocument.Element("Assembly").Element("Types").Elements("Type")
                                      where a.Attribute("Name").Value.Equals(itemField.Attribute("Value").Value)
                                      select a).FirstOrDefault();


                    XElement supportNode = memberNode.Element("SupportByLibrary");
                    string[] returnArray = new string[supportNode.Elements("Version").Count() + 1];
                    returnArray[0] = supportNode.Attribute("Name").Value;
                    int i = 1;
                    foreach (XElement item in supportNode.Elements("Version"))
                    {
                        returnArray[i] = item.Value;
                        i++;
                    }

                   return returnArray;
                }
            }

            throw new ArgumentException(itemField + " is unkown");
        }
        
        private static string[] GetSupportByLibraryFieldSet(XElement itemField, string libName)
        {
            switch (itemField.Name.LocalName)
            {
                case "FieldSet":
                {
                        XDocument apiDocument = GetDocument(libName);
                        XElement memberNode = (from a in apiDocument.Element("Assembly").Element("Enums").Elements("Enum").Elements("Member")
                                               where a.Attribute("Name").Value.Equals(itemField.FirstAttribute.Value)
                                               select a).FirstOrDefault();


                        if (null == memberNode)
                            memberNode = (from a in apiDocument.Element("Assembly").Element("Types").Elements("Type")
                                          where a.Attribute("Name").Value.Equals(itemField.Attribute("Value").Value)
                                          select a).FirstOrDefault();


                        XElement supportNode = memberNode.Element("SupportByLibrary");
                        string[] returnArray = new string[supportNode.Elements("Version").Count() + 1];
                        returnArray[0] = supportNode.Attribute("Name").Value;
                        int i = 1;
                        foreach (XElement item in supportNode.Elements("Version"))
                        {
                            returnArray[i] = item.Value;
                            i++;
                        }

                        return returnArray;
                 }
            }

            throw new ArgumentException(itemField + " is unkown");
        }

        private static string[] GetSupportByLibraryTypeEntity(string entityName, XElement itemType, XElement itemField, string libName)
        {
            switch (itemField.Name.LocalName)
            {
                case "Call":
                    {
                        
                        XDocument apiDocument = GetDocument(libName);

                        XElement memberNode = (from a in itemType.Elements("Method")
                                               where a.Attribute("Name").Value.Equals(entityName)
                                               select a).FirstOrDefault();

                        if(null==memberNode)
                            memberNode = (from a in itemType.Elements("Property")
                                          where a.Attribute("Name").Value.Equals(entityName)
                                          select a).FirstOrDefault();

                        XElement supportNode = memberNode.Element("SupportByLibrary");
                        string[] returnArray = new string[supportNode.Elements("Version").Count() + 1];
                        returnArray[0] = supportNode.Attribute("Name").Value;
                        int i = 1;
                        foreach (XElement item in supportNode.Elements("Version"))
                        {
                            returnArray[i] = item.Value;
                            i++;
                        }

                        return returnArray;
                    }
            }

            throw new ArgumentException(itemField + " is unkown");
        }

        private static string[] GetSupportByLibrary(XElement itemField, ref string libName)
        {
            switch (itemField.Name.LocalName)
            {
                case "Var":
                case "Field":
                case "Property":
                case "ReturnValue":
                {
                    string[] splitArray = itemField.Attribute("Type").Value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    XDocument apiDocument = GetDocument(splitArray[1]);

                    XElement memberNode = (from a in apiDocument.Element("Assembly").Element("Enums").Elements("Enum")
                                           where a.Attribute("Name").Value.Equals(splitArray[splitArray.Length - 1])
                                           select a).FirstOrDefault();

                    if(null == memberNode)
                        memberNode = (from a in apiDocument.Element("Assembly").Element("Types").Elements("Type")
                                      where a.Attribute("Name").Value.Equals(splitArray[splitArray.Length - 1])
                                      select a).FirstOrDefault();


                    XElement supportNode = memberNode.Element("SupportByLibrary");
                    string[] returnArray = new string[supportNode.Elements("Version").Count()+1];
                    returnArray[0] = supportNode.Attribute("Name").Value;
                    libName = supportNode.Attribute("Name").Value;
                    int i = 1;
                    foreach (XElement item in supportNode.Elements("Version"))
                    {
                        returnArray[i] = item.Value; 
                        i++;
                    }
                    
                    return returnArray;
                }
            }

            throw new ArgumentException(itemField + " is unkown");
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
            if (name.EndsWith("Api"))
                name = name.Substring(0, name.Length - 3);

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

        private static bool MethodBodyPassed(XElement itemClass, XElement itemMethod, AssemblyAnalyzerSettings settings, ref string[] warnings, bool isPropertyBody)
        {
            string entityType = "Method";
            if (isPropertyBody)
                entityType = "Property";

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
                            string warning = string.Format("class {0}: " + entityType + " {1}; ReturnValue: {4}; SupportByLibrary {3}, {2}",
                                                            itemClass.Attribute("Name").Value, itemMethod.Attribute("Name").Value,
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
                            string warning = string.Format("class {0}: " + entityType + " {1}; Variable: {4}; SupportByLibrary {3}, {2}",
                                                            itemClass.Attribute("Name").Value, itemMethod.Attribute("Name").Value,
                                                            ToString(libs), libName, itemEntity.Attribute("Type").Value);

                            listWarnings.Add(warning);
                        }

                        string type = itemEntity.Attribute("Type").Value;
                        foreach (XElement itemSet in itemEntity.Elements("Set"))
                        {
                            string setValue = itemSet.Attribute("Value").Value;
                            if (IsEnum(type))
                            {
                                itemSet.Attribute("Value").Value = GetEnumMemberName(type, setValue);
                                libs = GetSupportByLibrarySet(itemSet, libName);                                
                            }
                            if (!FieldPassed(libName, libs, settings))
                            {
                                string warning = string.Format("class {0}: " + entityType + " {1}; Variable Set: {4}; SupportByLibrary {3}, {2}",
                                                                itemClass.Attribute("Name").Value, itemMethod.Attribute("Name").Value,
                                                                ToString(libs), libName, itemSet.Attribute("Value").Value);

                                listWarnings.Add(warning);
                            }
                        }

                        break;
                    }
                    case "FieldSet":
                    {
                        string[] libs = null;
                        XElement field = (from a in itemClass.Element("Fields").Elements("Field")
                                          where a.Attribute("Name").Value.Equals(itemEntity.FirstAttribute.Name.LocalName)
                                   select a).FirstOrDefault();
                       
                        string type = field.Attribute("Type").Value;
                        string libName = (type.Split(new string[]{"."},StringSplitOptions.RemoveEmptyEntries))[1];

                        string setValue = itemEntity.FirstAttribute.Value;
                        if (IsEnum(type))
                        {
                            itemEntity.FirstAttribute.Value = GetEnumMemberName(type, setValue);
                            libs = GetSupportByLibraryFieldSet(itemEntity, libName);
                        }

                        if (!FieldPassed(libName, libs, settings))
                        {
                            string warning = string.Format("class {0}: " + entityType + " {1}; Field Set: {4}; SupportByLibrary {2}",
                                                            itemClass.Attribute("Name").Value, itemMethod.Attribute("Name").Value,
                                                            ToString(libs), libName, itemEntity.FirstAttribute.Value);

                            listWarnings.Add(warning);
                        }

                        break;
                    }
                    case "Call":
                    {
                        if (itemEntity.Attribute("Name").Value != ".ctor")
                        {
                            string targetType = "";
                            string name = "";
                            if (itemEntity.Attribute("Name").Value.StartsWith("set_") || itemEntity.Attribute("Name").Value.StartsWith("get_"))
                            {
                                name = itemEntity.Attribute("Name").Value.Substring(4);
                                targetType = "Property";
                            }
                            else
                            {
                                name = itemEntity.Attribute("Name").Value;
                                targetType = "Method";
                            }

    
                            string[] splitAray = itemEntity.Attribute("Type").Value.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);

                            XDocument doc = GetDocument(splitAray[1]);

                            int paramsCount = 0;
                            if ("Property" != targetType)
                                paramsCount = itemEntity.Elements("Param").Count();

                            XElement entity = GetEntityNode(doc, splitAray[splitAray.Length - 1], name, targetType, paramsCount);

                           
                            XElement type = (from a in doc.Element("Assembly").Elements("Types").Elements("Type")
                                             where a.Attribute("Name").Value.Equals(splitAray[splitAray.Length-1])
                                              select a).FirstOrDefault();
                            /*
                           XElement entity = (from a in type.Elements(targetType)
                                              where a.Attribute("Name").Value.Equals(name)
                                             select a).FirstOrDefault();
                           */

                            string[] libs = null;
                            libs = GetSupportByLibraryTypeEntity(name, type, itemEntity, splitAray[1]);
                            if (!FieldPassed(splitAray[1], libs, settings))
                            {
                                string warning = string.Format("class {0}: " + entityType + " {1}; Call: {4}.{5}; SupportByLibrary {2}",
                                                                itemClass.Attribute("Name").Value, itemMethod.Attribute("Name").Value,
                                                                ToString(libs), splitAray[1], itemEntity.FirstAttribute.Value, name);

                                listWarnings.Add(warning);
                            }

                            foreach (XElement itemParam in itemEntity.Elements("Param"))
                            {
                                
                            }

                        }
                        break;
                    }
                }
            }

            warnings = new string[listWarnings.Count];
            int i = 0;
            foreach (string item in listWarnings)
            {
                warnings[i] = item;
                i++;
            }

            return warnings.Length > 0 ? false : true;
        }

        private static XElement GetEntityNode(XDocument doc, string typeName, string entityName, string targetType, int paramsCount)
        {
            XElement type = (from a in doc.Element("Assembly").Elements("Types").Elements("Type")
                             where a.Attribute("Name").Value.Equals(typeName)
                             select a).FirstOrDefault();

            var entities = (from a in type.Elements(targetType)
                               where a.Attribute("Name").Value.Equals(entityName)
                               select a);

            foreach (var item in entities)
            {
                if (item.Elements("Param").Count() == paramsCount)
                    return item;
            }

            throw new ArgumentException("Entity not found");
        }

        private static string Space(string space, int count)
        {
            string result = "";
            for (int i = 1; i <= count; i++)
                result += space;
            return result;
        }

        #endregion
    }
}
