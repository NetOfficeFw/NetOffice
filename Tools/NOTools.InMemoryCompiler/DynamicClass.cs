using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace NOTools.InMemoryCompiler
{
    /// <summary>
    /// DynamicAssembly class definition
    /// </summary>
    public class DynamicClass
    {
        #region Ctor

        /// <summary>s
        /// Creates an instance of the class
        /// </summary>
        internal DynamicClass(DynamicAssembly parent)
        {
            Parent = parent;
            Usings = new DynamicUsingsCollection();
            Interfaces = new DynamicInterfaceCollection();
            Properties = new DynamicPropertyCollection();
            Methods = new DynamicMethodCollection();            
        }

        #endregion
        
        #region Properties

        /// <summary>
        /// Parent assembly definition
        /// </summary>
        public DynamicAssembly Parent { get; private set; }

        /// <summary>
        /// Name of the class
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Namespace of the class. (AssemblyName by default)
        /// </summary>
        public string Namespace { get; set; }

        /// <summary>
        /// Access/Visibility of the class
        /// </summary>
        public AccessModifier Modifier { get; set; }

        /// <summary>
        /// Usings(Imports in VB) "System.Data" for example
        /// </summary>
        public DynamicUsingsCollection Usings { get; private set; }

        /// <summary>
        /// Interface implementations
        /// </summary>
        public DynamicInterfaceCollection Interfaces { get; private set; }

        /// <summary>
        /// Methods of the class
        /// </summary>
        public DynamicMethodCollection Methods { get; private set; }

        /// <summary>
        /// Properties of the class
        /// </summary>
        public DynamicPropertyCollection Properties { get; private set; }

        /// <summary>
        /// SourceCode template
        /// </summary>
        internal static string Template
        {
            get
            {
                if (null == _template)
                    _template = ReadRessourceTextFile("NetOfficeDeveloperAddin.Compiler.ClassTemplate.txt");
                return _template;
            }
        }
        private static string _template;

        #endregion

        #region Static Helper Methods to generate the class

        /// <summary>
        /// Generate all usings from DynamicClass instance in a code template
        /// </summary>
        /// <param name="dynamicClass">DynamicClass definition</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddUsingsToTemplate(DynamicClass dynamicClass, string template)
        {
            StringBuilder stringBuilder = new StringBuilder();
            foreach (var item in dynamicClass.Usings)
                stringBuilder.Append(String.Format("using {0};{1}", item, Environment.NewLine));
            return template.Replace("%Usings%", stringBuilder.ToString());
        }

        // <summary>
        /// Generate the namespace from DynamicClass instance in a code template
        /// </summary>
        /// <param name="classNamespace">namespace to set</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddNamespaceToTemplate(string classNamespace, string template)
        {
            return template.Replace("%Namespace%", classNamespace);
        }

        /// <summary>
        /// Generate the name from DynamicClass instance in a code template
        /// </summary>
        /// <param name="dynamicClass">DynamicClass definition</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddNameToTemplate(DynamicClass dynamicClass, string template)
        {
            return template.Replace("%Name%", dynamicClass.Name);
        }

        /// <summary>
        /// Generate all interface implementations from DynamicClass instance in a code template
        /// </summary>
        /// <param name="dynamicClass">DynamicClass definition</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddInterfacesToTemplate(DynamicClass dynamicClass, string template)
        {
            if (dynamicClass.Interfaces.Count == 0)
                return template.Replace("%Interfaces%", "");

            string interfaceString = ":";
            foreach (var item in dynamicClass.Interfaces)
                interfaceString += item.Name + ",";

            interfaceString = interfaceString.Substring(0, interfaceString.Length - 1);

            return template.Replace("%Interfaces%", interfaceString);
        }

        /// <summary>
        /// Generate all properties from DynamicClass instance in a code template
        /// </summary>
        /// <param name="dynamicClass">DynamicClass definition</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddPropertiesToTemplate(DynamicClass dynamicClass, string template)
        {
            StringBuilder stringBuilderProperties = new StringBuilder();
            StringBuilder stringBuilderParams = new StringBuilder();
            StringBuilder stringBuilderSet = new StringBuilder();
            StringBuilder stringBuilderSet2 = new StringBuilder();

            int i = 0;
            foreach (var item in dynamicClass.Properties)
            {
                if (i > 0)
                {
                    stringBuilderParams.Append(",");
                    stringBuilderSet2.Append(",");
                }

                string[] splitArray = item.ToString().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                stringBuilderProperties.Append(String.Format("\t\tinternal {0} {1}{{get; set;}}{2}", splitArray[0], splitArray[1], Environment.NewLine));
                stringBuilderParams.Append(String.Format("{0} param{1}", splitArray[0], i));
                stringBuilderSet.Append(String.Format("\t\t\t{0}=param{1};{2}", splitArray[1], i, Environment.NewLine));
                stringBuilderSet2.Append(String.Format("param{0}", i));
                i++;
            }

            template = template.Replace("%Properties%", stringBuilderProperties.ToString());
            template = template.Replace("%InitializeParams%", stringBuilderParams.ToString());
            template = template.Replace("%ParamsToProperties%", stringBuilderSet.ToString());
            template = template.Replace("%InitializeParams2%", stringBuilderSet2.ToString());
            return template;
        }

        /// <summary>
        /// Generate all methods from DynamicClass instance in a code template
        /// </summary>
        /// <param name="dynamicClass">DynamicClass definition</param>
        /// <param name="template">Source code template</param>
        /// <returns>Updated code template</returns>
        internal static string AddMethodsToTemplate(DynamicClass dynamicClass, string template)
        {
            StringBuilder stringBuilderMethods = new StringBuilder();

            foreach (var item in dynamicClass.Methods)
            {
                string methodTemplate = CreateMethodTemplate(item);
                methodTemplate = methodTemplate.Replace("%Code%", item.MethodCode);
                stringBuilderMethods.Append(methodTemplate);
            }

            return template.Replace("%Methods%", stringBuilderMethods.ToString());
        }
        
        /// <summary>
        /// Create an empty method block from DynamicMethod instance with "%Code%" instead of the methodCode
        /// </summary>
        /// <param name="method">Method definition</param>
        /// <returns>Method template</returns>
        internal static string CreateMethodTemplate(DynamicMethod method)
        {
            string result = String.Format("\t\tpublic {0} {1}(", method.ReturnValue != "" ? method.ReturnValue : "void", method.Name);

            for (int i = 0; i < method.Parameters.Count; i++)
            {
                string[] splitArray = method.Parameters[i].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                if (splitArray.Length == 2)
                    result += String.Format("{0} {1}", splitArray[0].Trim(), splitArray[1].Trim());
                else if (splitArray.Length == 3)
                    result += String.Format("{0} {1} {2}", splitArray[0].Trim(), splitArray[1].Trim(), splitArray[2].Trim());

                if (i < method.Parameters.Count - 1)
                    result += ",";
            }

            result += ")" + Environment.NewLine + "\t\t{" + Environment.NewLine + "%Code%" + Environment.NewLine + "\t\t}" + Environment.NewLine;

            return result;
        }

        /// <summary>
        /// Read ans embeddded textfile into a string
        /// </summary>
        /// <param name="fullQualifiedPath">Ressource location</param>
        /// <returns>Textfile content</returns>
        private static string ReadRessourceTextFile(string fullQualifiedPath)
        {
            System.IO.Stream ressourceStream;
            System.IO.StreamReader textStreamReader;
            try
            {
                ressourceStream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(fullQualifiedPath);
                if (ressourceStream == null)
                    throw (new System.IO.IOException("Error accessing resource Stream."));

                textStreamReader = new System.IO.StreamReader(ressourceStream);
                if (textStreamReader == null)
                    throw (new System.IO.IOException("Error accessing resource File."));

                string text = textStreamReader.ReadToEnd();
                ressourceStream.Close();
                textStreamReader.Close();
                return text;
            }
            catch (Exception exception)
            {
                throw (exception);
            }
        }

        #endregion

        #region Overrides

        /// <summary>
        /// Returns a string that represents the instance
        /// </summary>
        /// <returns>System.String</returns>
        public override string ToString()
        {            
            string code = DynamicClass.Template;
            code = DynamicClass.AddUsingsToTemplate(this,code);
            code = DynamicClass.AddInterfacesToTemplate(this, code);
            code = DynamicClass.AddNamespaceToTemplate(String.IsNullOrWhiteSpace(Namespace) ? Parent.Name : Namespace, code);
            code = DynamicClass.AddNameToTemplate(this, code);
            code = DynamicClass.AddPropertiesToTemplate(this, code);
            code = DynamicClass.AddMethodsToTemplate(this, code);
            return code;
        }

        #endregion
    }
}
