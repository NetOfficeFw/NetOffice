using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.IO.Compression;
using Mono.Cecil;
using Mono.Cecil.Cil;
using Project = NetOffice.MSProjectApi;
using Visio = NetOffice.VisioApi;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeCompatibility
{
    class NetOfficeSupportTable
    {
        AssemblyDefinition _assOffice;
        AssemblyDefinition _assExcel;
        AssemblyDefinition _assWord;
        AssemblyDefinition _assOutlook;
        AssemblyDefinition _assPowerPoint;
        AssemblyDefinition _assAccess;
        AssemblyDefinition _assMSProject;
        AssemblyDefinition _assVisio;
        AssemblyDefinition _assPublisher;

        Assembly _thisAssembly = Assembly.GetExecutingAssembly();
        
        public NetOfficeSupportTable()
        {
            AssemblyName[] referencedAssemblies = _thisAssembly.GetReferencedAssemblies();
            foreach (AssemblyName item in referencedAssemblies)
            {
                {
                    if (item.Name.StartsWith("OfficeApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assOffice = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("ExcelApi"))                    
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assExcel = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("WordApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assWord = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("OutlookApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assOutlook = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("OutlookApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assOutlook = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("PowerPointApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assPowerPoint = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("AccessApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assAccess = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("MSProjectApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assMSProject = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("VisioApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assVisio = AssemblyDefinition.ReadAssembly(stream);
                    }
                    else if (item.Name.StartsWith("PublisherApi"))
                    {
                        string assemblyPath = GetPhysicalPath(item);
                        Stream stream = System.IO.File.OpenRead(assemblyPath);
                        _assPublisher = AssemblyDefinition.ReadAssembly(stream);
                    }
                }
            }        
        }

        /// <summary>
        /// returns enum member name for an enum value
        /// </summary>
        /// <param name="fullQualifiedName">Name of enum</param>
        /// <param name="value">target value</param>
        /// <returns></returns>
        public string GetEnumMemberNameFromValue(string fullQualifiedName, int value)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetName(fullQualifiedName);

            AssemblyDefinition assembly = GetAssembly(library);
            if (null == assembly)
                return null;

            string fullQualifiedTypeName = GetQualifiedTypeCallType(fullQualifiedName);

            TypeDefinition typeDef = (from a in assembly.Modules[0].Types where a.FullName.Equals(fullQualifiedTypeName, StringComparison.InvariantCultureIgnoreCase) select a).FirstOrDefault();
            if (null == typeDef)
                return null;

            FieldDefinition fieldDef = (from a in typeDef.Fields where value.Equals(a.Constant) select a).FirstOrDefault();
            if (null == fieldDef)
                return null;

            return fieldDef.Name;
        }

        /// <summary>
        /// returns a string array with supported versions for an enum
        /// </summary>
        /// <param name="fullQualifiedName"></param>
        /// <param name="value"></param>
        /// <returns></returns>
        public string[] GetEnumMemberSupport(string fullQualifiedName, int value)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetName(fullQualifiedName);

            AssemblyDefinition assembly = GetAssembly(library);
            if (null == assembly)
                return null;

            string fullQualifiedTypeName = GetQualifiedTypeCallType(fullQualifiedName);

            TypeDefinition typeDef = (from a in assembly.Modules[0].Types where a.FullName.Equals(fullQualifiedTypeName, StringComparison.InvariantCultureIgnoreCase) select a).FirstOrDefault();
            if (null == typeDef)
                return null;

            FieldDefinition fieldDef = (from a in typeDef.Fields where value.Equals(a.Constant) select a).FirstOrDefault();
            if (null == fieldDef)
                return null;

            CustomAttribute typeDefAttribute = (from a in fieldDef.CustomAttributes
                                                where a.AttributeType.FullName.Equals("NetOffice.SupportByLibraryAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.Attributes.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                select a).FirstOrDefault();
            if (null == typeDefAttribute)
                return null;

            CustomAttributeArgument[] versions = typeDefAttribute.ConstructorArguments[1].Value as CustomAttributeArgument[];
            string[] result = new string[versions.Length];
            for (int i = 0; i < versions.Length; i++)
                result[i] = Convert.ToString(versions[i].Value);
            return result;
        }

        /// <summary>
        ///  returns a string array with supported versions for a method or property
        /// </summary>
        /// <param name="fullQualifiedName"></param>
        /// <returns></returns>
        public string[] GetTypeCallSupport(string fullQualifiedName)
        {
            string library = GetLibrary(fullQualifiedName);
            string typeName = GetTypeName(fullQualifiedName);
            string methodName = GetName(fullQualifiedName);
            string[] parameters = GetParameters(fullQualifiedName);

            AssemblyDefinition assembly = GetAssembly(library);
            if (null == assembly)
                return null;

            string fullQualifiedTypeName = GetQualifiedTypeCallType(fullQualifiedName);

            TypeDefinition typeDef = (from a in assembly.Modules[0].Types where a.FullName.Equals(fullQualifiedTypeName, StringComparison.InvariantCultureIgnoreCase) select a).FirstOrDefault();
            if (null == typeDef)
                return null;

            string[] result = GetTypeCallSupportProperty(typeDef, methodName, parameters.Length);
            if (null == result)
                result = GetTypeCallSupportMethod(typeDef, methodName, parameters.Length);
            if (null == result)
                result = GetTypeCallSupportEvent(typeDef, methodName);

            return result;
        }

        /// <summary>
        /// returns a string array with supported versions for a type
        /// </summary>
        /// <param name="fullQualifiedName"></param>
        /// <returns></returns>
        public string[] GetTypeSupport(string fullQualifiedName)
        {
            if (fullQualifiedName.EndsWith("[]", StringComparison.InvariantCultureIgnoreCase))
                fullQualifiedName = fullQualifiedName.Substring(0, fullQualifiedName.Length - 2);

            string library = GetLibrary(fullQualifiedName);

            AssemblyDefinition assembly = GetAssembly(library);
            if (null == assembly)
                return null;

            TypeDefinition typeDef = (from a in assembly.Modules[0].Types where a.FullName.Equals(fullQualifiedName, StringComparison.InvariantCultureIgnoreCase) select a).FirstOrDefault();
            if (null == typeDef)
                return null;
            CustomAttribute typeDefAttribute = (from a in typeDef.CustomAttributes
                                                where a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.Attributes.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                select a).FirstOrDefault();
            if (null == typeDefAttribute)
                return null;

            CustomAttributeArgument[] versions = typeDefAttribute.ConstructorArguments[1].Value as CustomAttributeArgument[];
            string[] result = new string[versions.Length];
            for (int i = 0; i < versions.Length; i++)
                result[i] = Convert.ToString(versions[i].Value);
            return result;
        }
        
        /// <summary>
        /// Gets an  AssemblyDefinition
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private AssemblyDefinition GetAssembly(string name)
        {
            switch (name)
            {
                case "Office":
                    return _assOffice;
                case "Excel":
                    return _assExcel;
                case "Word":
                    return _assWord;
                case "Outlook":
                    return _assOutlook;
                case "PowerPoint":
                    return _assPowerPoint;
                case "Access":
                    return _assAccess;
                case "MSProject":
                    return _assMSProject;
                case "Visio":
                    return _assVisio;
                case "Publisher":
                    return _assPublisher;
                default:
                    return null;
            }
        }

        /// <summary>
        /// returns the containing library name of the qualifier
        /// </summary>
        /// <param name="fullQualifiedName"></param>
        /// <returns></returns>
        public static string GetLibrary(string fullQualifiedName)
        {
            string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
            if (array == null || array.Length < 2 || array[1].Length < 3)
                return null;
            string documentName = array[1].Substring(0, array[1].Length - 3);
            return documentName;
        }

        public static string GetQualifiedTypeCallType(string fullQualifiedName)
        {
            string[] array = fullQualifiedName.Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
            return array[0];
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
            if (fullQualifiedName.Contains(".Native."))
            {
                string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                return array[3].Substring(0, array[3].IndexOf("::", StringComparison.InvariantCultureIgnoreCase));
            }
            else
            {
                string[] array = fullQualifiedName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                return array[2].Substring(0, array[2].IndexOf("::", StringComparison.InvariantCultureIgnoreCase));
            }
        }

        /// <summary>
        /// returns the name of the qualifier
        /// </summary>
        /// <param name="fullQualifiedName"></param>
        /// <returns></returns>
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


        private string[] GetTypeCallSupportProperty(TypeDefinition typeDef, string methodName, int parametersCount)
        {
            int targetParamsCount = parametersCount;
            if (methodName.StartsWith("set_"))
                targetParamsCount -= 1;
            if (methodName.StartsWith("set_") || methodName.StartsWith("get_"))
                methodName = methodName.Substring("get_".Length);
            PropertyDefinition targetProperty = (from a in typeDef.Properties where a.Name.Equals(methodName, StringComparison.InvariantCultureIgnoreCase) && a.Parameters.Count() == targetParamsCount select a).FirstOrDefault();
            if (null == targetProperty)
                return null;

            CustomAttribute typeDefAttribute = (from a in targetProperty.CustomAttributes
                                                where a.AttributeType.FullName.Equals("NetOffice.SupportByLibraryAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.Attributes.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                select a).FirstOrDefault();
            if (null == typeDefAttribute)
                return null;

            CustomAttributeArgument[] versions = typeDefAttribute.ConstructorArguments[1].Value as CustomAttributeArgument[];
            string[] result = new string[versions.Length];
            for (int i = 0; i < versions.Length; i++)
                result[i] = Convert.ToString(versions[i].Value);
            return result;
        }

        private string[] GetTypeCallSupportMethod(TypeDefinition typeDef, string methodName, int parametersCount)
        {
            MethodDefinition targetMethod = (from a in typeDef.Methods where a.Name.Equals(methodName, StringComparison.InvariantCultureIgnoreCase) && a.Parameters.Count() == parametersCount select a).FirstOrDefault();
            if (null == targetMethod)
                return null;

            CustomAttribute typeDefAttribute = (from a in targetMethod.CustomAttributes
                                                where a.AttributeType.FullName.Equals("NetOffice.SupportByLibraryAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.Attributes.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                select a).FirstOrDefault();
            if (null == typeDefAttribute)
                return null;

            CustomAttributeArgument[] versions = typeDefAttribute.ConstructorArguments[1].Value as CustomAttributeArgument[];
            string[] result = new string[versions.Length];
            for (int i = 0; i < versions.Length; i++)
                result[i] = Convert.ToString(versions[i].Value);
            return result;
        }

        private string[] GetTypeCallSupportEvent(TypeDefinition typeDef, string methodName)
        {

            if (methodName.StartsWith("add_"))
                methodName = methodName.Substring("add_".Length);
            if (methodName.StartsWith("remove_"))
                methodName = methodName.Substring("remove_".Length);

            EventDefinition targetEvent = (from a in typeDef.Events where a.Name.Equals(methodName, StringComparison.InvariantCultureIgnoreCase) select a).FirstOrDefault();
            if (null == targetEvent)
                return null;

            CustomAttribute typeDefAttribute = (from a in targetEvent.CustomAttributes
                                                where a.AttributeType.FullName.Equals("NetOffice.SupportByLibraryAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                    || a.AttributeType.FullName.Equals("NetOffice.Attributes.SupportByVersionAttribute", StringComparison.InvariantCultureIgnoreCase)
                                                select a).FirstOrDefault();
            if (null == typeDefAttribute)
                return null;

            CustomAttributeArgument[] versions = typeDefAttribute.ConstructorArguments[1].Value as CustomAttributeArgument[];
            string[] result = new string[versions.Length];
            for (int i = 0; i < versions.Length; i++)
                result[i] = Convert.ToString(versions[i].Value);
            return result;
        }

        private string GetPhysicalPath(AssemblyName assemblyName)
        {
            string directoryName = Program.DependencySubFolder;
            string fileName = assemblyName.Name;
            if (fileName.IndexOf(",") > -1)
                fileName = assemblyName.Name.Substring(0, assemblyName.Name.IndexOf(","));
            string fullFileName = System.IO.Path.Combine(directoryName, fileName + ".dll");

            if (!File.Exists(fullFileName))
            {
                directoryName = Program.DependencyReleaseSubFolder;
                fileName = assemblyName.Name;
                if (fileName.IndexOf(",") > -1)
                    fileName = assemblyName.Name.Substring(0, assemblyName.Name.IndexOf(","));
                fullFileName = System.IO.Path.Combine(directoryName, fileName + ".dll");
            }
            
            return fullFileName;
        }

        private static Stream ReadEmbeddedAssembly(string ressourcePath)
        {
            System.IO.Stream ressourceStream = null;
            string assemblyName = System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
            ressourcePath = assemblyName + ".OfficeCompatibility." + ressourcePath;
            ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(ressourcePath);
            if (null == ressourceStream)
                throw new System.IO.FileLoadException(ressourcePath + " not found");
            System.IO.MemoryStream outStream = new System.IO.MemoryStream();
            using (GZipStream Decompress = new GZipStream(ressourceStream, CompressionMode.Decompress))
            {
                Decompress.CopyTo(outStream);
            }
            outStream.Seek(0, SeekOrigin.Begin);
            ressourceStream.Close();
            return outStream;
        }
    }
}