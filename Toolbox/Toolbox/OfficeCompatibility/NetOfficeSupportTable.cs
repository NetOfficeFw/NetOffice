using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text;
using System.IO;
using System.IO.Compression;


namespace NetOffice.DeveloperToolbox.OfficeCompatibility
{
    class NetOfficeSupportTable
    {
        XDocument _office;
        XDocument _excel;
        XDocument _word;
        XDocument _outlook;
        XDocument _powerPoint;
        XDocument _access;

        public XDocument Decompress(System.IO.Stream ressourceStream)
        {
            System.IO.MemoryStream outStream = new System.IO.MemoryStream(); 
            using (GZipStream Decompress = new GZipStream(ressourceStream, CompressionMode.Decompress))
            {
                Decompress.CopyTo(outStream);
            }
            outStream.Seek(0, SeekOrigin.Begin);
            XDocument document = XDocument.Load(outStream);
            return document;
        }

        private static Stream ReadStreamFromRessource(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".OfficeCompatibility." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (null == ressourceStream)
                throw new System.IO.FileLoadException(ressourcePath + " not found");
            return ressourceStream;

        }

        public NetOfficeSupportTable()
        {
            Stream stream = ReadStreamFromRessource("NetOfficeDocuFiles.Office.Description.xml.gz");
            _office = Decompress(stream);

            stream = ReadStreamFromRessource("NetOfficeDocuFiles.Excel.Description.xml.gz");
            _excel = Decompress(stream);
           
            stream = ReadStreamFromRessource("NetOfficeDocuFiles.Word.Description.xml.gz");
            _word = Decompress(stream);
           
            stream = ReadStreamFromRessource("NetOfficeDocuFiles.Outlook.Description.xml.gz");
            _outlook = Decompress(stream);

            stream = ReadStreamFromRessource("NetOfficeDocuFiles.PowerPoint.Description.xml.gz");
            _powerPoint = Decompress(stream);

            stream = ReadStreamFromRessource("NetOfficeDocuFiles.Access.Description.xml.gz");
            _access = Decompress(stream);
        }

        public string GetEnumMemberNameFromValue(string fullQualifiedName, int value)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetName(fullQualifiedName);
            XDocument document = GetDocument(library);
            if (null == document)
                return null;

            XElement enumNode = (from a in document.Element("Assembly").Element("Enums").Elements("Enum")
                                 where a.Attribute("Name").Value.Equals(typeName)
                                 select a).FirstOrDefault();
            if (null == enumNode)
                return null;

            XElement memberNode = (from a in enumNode.Element("Members").Elements("Member")
                                   where a.Attribute("Value").Value.Equals(value.ToString())
                                   select a).FirstOrDefault();
            if (null == memberNode)
                return null;

            return memberNode.Attribute("Name").Value;
        }

        public string[] GetEnumMemberSupport(string fullQualifiedName, int value)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetName(fullQualifiedName);
            XDocument document = GetDocument(library);
            if (null == document)
                return null;

            XElement enumNode = (from a in document.Element("Assembly").Element("Enums").Elements("Enum")
                                 where a.Attribute("Name").Value.Equals(typeName)
                                 select a).FirstOrDefault();
            if (null == enumNode)
                return null;

            XElement memberNode = (from a in enumNode.Element("Members").Elements("Member")
                                   where a.Attribute("Value").Value.Equals(value.ToString())
                                 select a).FirstOrDefault();
            if (null == memberNode)
                return null;

            string[] result = new string[memberNode.Element("SupportByLibrary").Elements("Version").Count()];
            int i = 0;
            foreach (XElement item in memberNode.Element("SupportByLibrary").Elements("Version"))
            {
                result[i] = item.Value;
                i++;
            }
            return result;
        }

        public string[] GetTypeCallSupport(string fullQualifiedName)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetTypeName(fullQualifiedName);
            string methodName = GetName(fullQualifiedName);
            string[] parameters = GetParameters(fullQualifiedName);
            XDocument document = GetDocument(library);
            if (null == document)
                return null;

            XElement typeNode = (from a in document.Element("Assembly").Element("Types").Elements("Type")
                                 where a.Attribute("Name").Value.Equals(typeName)
                                 select a).FirstOrDefault();

            XElement methodNode = null;
            if (typeNode.Element("Methods") != null)
            {
                methodNode = (from a in typeNode.Element("Methods").Elements("Method")
                              where a.Attribute("Name").Value.Equals(methodName)
                              select a).FirstOrDefault();

                foreach (XElement itemParameters in methodNode.Elements("Parameters"))
                {
                    int count = itemParameters.Elements("Parameter").Count();
                    if (count == parameters.Count())
                    {
                        string[] result = new string[itemParameters.Element("SupportByLibrary").Elements("Version").Count()];
                        int i = 0;
                        foreach (XElement item in itemParameters.Element("SupportByLibrary").Elements("Version"))
                        {
                            result[i] = item.Value;
                            i++;
                        }
                        return result;
                    }
                }

            }
            else
            {
                methodName = methodName.Substring(0, methodName.Length - 5);
                methodNode = (from a in typeNode.Element("Events").Elements("Event")
                              where a.Attribute("Name").Value.Equals(methodName)
                              select a).FirstOrDefault();

                string[] result = new string[methodNode.Element("SupportByLibrary").Elements("Version").Count()];
                int i = 0;
                foreach (XElement item in methodNode.Element("SupportByLibrary").Elements("Version"))
                {
                    result[i] = item.Value;
                    i++;
                }
                return result;

            } 

            return null;
        }

        public string[] GetTypeSupport(string fullQualifiedName)
        {
            if (fullQualifiedName.EndsWith("[]", StringComparison.InvariantCultureIgnoreCase))
                fullQualifiedName = fullQualifiedName.Substring(0, fullQualifiedName.Length - 2);

            string library = GetLibrary(fullQualifiedName);
            string name = GetName(fullQualifiedName);
            XDocument document = GetDocument(library);
            if (null == document)
                return null;

            XElement typeNode = (from a in document.Element("Assembly").Element("Types").Elements("Type")
                                 where a.Attribute("Name").Value.Equals(name) 
                                 select a).FirstOrDefault();

            if (null == typeNode)
            {
                typeNode = (from a in document.Element("Assembly").Element("Enums").Elements("Enum")
                            where a.Attribute("Name").Value.Equals(name)
                            select a).FirstOrDefault();
            }

            if (null == typeNode)
                return null;

            string[] result = new string[typeNode.Element("SupportByLibrary").Elements("Version").Count()];
            int i=0;
            foreach (XElement item in typeNode.Element("SupportByLibrary").Elements("Version"))
	        {
                result[i] = item.Value;
                i++;
	        }

            return result;
        }

        public static string GetLibrary(string fullQualifiedName)
        {
            string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            if (array[1].Length < 3)
                return null;
            string documentName = array[1].Substring(0, array[1].Length - 3);
            return documentName;
        }

        public static string[] GetParameters(string fullQualifiedName)
        {
            string[] array = fullQualifiedName.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
            string part = array[array.Length - 1];
            part = part.Substring(part.IndexOf("(", StringComparison.InvariantCultureIgnoreCase));
            array = part.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i < array.Length; i++)
                array[i] = array[i].Replace("(", "").Replace(")", "");

            List<string> validateList = new List<string>();
            foreach (string item in array)
            {
                if (!string.IsNullOrEmpty(item))
                    validateList.Add(item);
            }
            return validateList.ToArray();
        }

        public static string GetTypeName(string fullQualifiedName)
        {
            string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            return array[2].Substring(0, array[2].IndexOf("::",StringComparison.InvariantCultureIgnoreCase));
        }

        public static string GetName(string fullQualifiedName)
        {
            if (fullQualifiedName.IndexOf("(", StringComparison.InvariantCultureIgnoreCase) > -1)
            {
                string[] array = fullQualifiedName.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
                string part = array[array.Length - 1];
                part = part.Substring(0, part.IndexOf("(", StringComparison.InvariantCultureIgnoreCase));
                return part;
            }
            else
            {
                string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                string documentName = array[array.Length - 1];
                if (documentName.EndsWith("[]", StringComparison.InvariantCultureIgnoreCase))
                    documentName = documentName.Substring(0, documentName.Length - 2);
                return documentName;
            }
           
        }

        private XDocument GetDocument(string name)
        {            
            switch (name)
            {
                case "Office":
                    return _office;
                case "Excel":
                    return _excel;
                case "Word":
                    return _word;
                case "Outlook":
                    return _outlook;
                case "PowerPoint":
                    return _powerPoint;
                case "Access":
                    return _access;
                default:
                    return null;
            }

        }
    }
}
