using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Mono.Cecil;
using Mono.Cecil.Cil;

namespace NetOffice.DeveloperToolbox.OfficeCompatibility
{
    public enum SupportVersion
    {
        NotUse = 0,
        Support = 1,
        NotSupport
    }

    public class SupportInfo
    {
        #region Fields

        SupportVersion _support;
        int _version;
        string _name;

        #endregion

        #region Construction
        
        internal SupportInfo(SupportVersion support, string name, int version)
        {
           _support = support;
           _name = name;
           _version = version;
        }

        #endregion

        #region Properties

        public SupportVersion Support
        {
           get
           {
               return _support;
           }
        }

        public int Version 
        {
           get
           {
               return _version;
           }
        }

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

        bool _containsNetOfficeReferences;
        XDocument _report;

        SupportInfo[] _office;
        SupportInfo[] _excel;
        SupportInfo[] _word;
        SupportInfo[] _outlook;
        SupportInfo[] _powerPoint;
        SupportInfo[] _access;
          
        #endregion

        #region Construction

        internal AnalyzerResult(bool containsNetOfficeReferences)
        {
            _containsNetOfficeReferences = containsNetOfficeReferences;
        }

        internal AnalyzerResult(XDocument report)
        {
            _report = report;
            _containsNetOfficeReferences = true;
            _office = new SupportInfo[5];
            _excel = new SupportInfo[5];
            _word = new SupportInfo[5];
            _outlook = new SupportInfo[5];
            _powerPoint = new SupportInfo[5];
            _access = new SupportInfo[5];

            RemoveDelegateTypes();

            SetupSupportInfo(_office, "Office");
            SetupSupportInfo(_excel, "Excel");
            SetupSupportInfo(_word, "Word");
            SetupSupportInfo(_outlook, "Outlook");
            SetupSupportInfo(_powerPoint, "PowerPoint");
            SetupSupportInfo(_access, "Access");
        }

        #endregion
            
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

        private void SetupSupportInfo(SupportInfo[] info, string name)
        {
            bool found09 = true;
            bool found10 = true;
            bool found11 = true;
            bool found12 = true;
            bool found14 = true;

            foreach (XElement item in _report.Element("Document").Element("Assembly").Element("Classes").Elements("Class"))
            {
                /*
                 var supportNodes = item.Descendants("SupportByLibrary");
                */

                var supportNodes = (from a in item.Descendants("SupportByLibrary")
                                    where a.Attribute("Api").Value.Equals(name, StringComparison.InvariantCultureIgnoreCase)
                                    select a);
                if ((null == supportNodes) || (supportNodes.Count() == 0))
                {
                    info[0] = new SupportInfo(SupportVersion.NotUse, name, 9);
                    info[1] = new SupportInfo(SupportVersion.NotUse, name, 10);
                    info[2] = new SupportInfo(SupportVersion.NotUse, name, 11);
                    info[3] = new SupportInfo(SupportVersion.NotUse, name, 12);
                    info[4] = new SupportInfo(SupportVersion.NotUse, name, 14);
                    return;
                }

                bool has09Support = false;
                bool has10Support = false;
                bool has11Support = false;
                bool has12Support = false;
                bool has14Support = false;

                foreach (XElement typeNodeItem in supportNodes)
                {
                    has09Support = IncludesVersion(typeNodeItem, "9");
                    has10Support = IncludesVersion(typeNodeItem, "10");
                    has11Support = IncludesVersion(typeNodeItem, "11");
                    has12Support = IncludesVersion(typeNodeItem, "12");
                    has14Support = IncludesVersion(typeNodeItem, "14");

                    if (!has09Support)
                        found09 = false;
                    if (!has10Support)
                        found10 = false;
                    if (!has11Support)
                        found11 = false;
                    if (!has12Support)
                        found12 = false;
                    if (!has14Support)
                        found14 = false;
                }
            }

            info[0] = new SupportInfo(BoolToSupportVersion(found09), name, 9);
            info[1] = new SupportInfo(BoolToSupportVersion(found10), name, 10);
            info[2] = new SupportInfo(BoolToSupportVersion(found11), name, 11);
            info[3] = new SupportInfo(BoolToSupportVersion(found12), name, 12);
            info[4] = new SupportInfo(BoolToSupportVersion(found14), name, 14);
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

        private static SupportVersion BoolToSupportVersion(bool value)
        {
            if (true == value)
                return SupportVersion.Support;
            else
                return SupportVersion.NotSupport;
        }

        public SupportInfo[] Office 
        {
            get             
            {
                return _office;
            }
        }

        public SupportInfo[] Excel
        {
            get
            {
                return _excel;
            }
        }

        public SupportInfo[] Word
        {
            get
            {
                return _word;
            }
        }

        public SupportInfo[] Outlook
        {
            get
            {
                return _outlook;
            }
        }

        public SupportInfo[] PowerPoint
        {
            get
            {
                return _powerPoint;
            }
        }

        public SupportInfo[] Access
        {
            get
            {
                return _access;
            }
        }
        

        #region Properties

        public XDocument Report 
        {
            get
            {
                return _report;
            }
        }

        public bool ContainsNetOfficeReferences 
        {
            get 
            {
                return _containsNetOfficeReferences;
            }
        }

        #endregion

    }
     
    public static class AssemblyAnalyzer
    {
        private static string _apiName = "NetOffice";

        private static NetOfficeSupportTable _netOfficeSupportTable = new NetOfficeSupportTable();

        #region Public Methods

        /// <summary>
        /// Analzye an assembly for NetOffice calls
        /// </summary>
        /// <param name="fullFileName"></param>
        /// <param name="analyzeDependencies"></param>
        public static AnalyzerResult AnalyzeAssembly(AssemblyDefinition assemblyDefinition)
        {
            if (!CheckNetOfficeReferencesExists(assemblyDefinition))
                return new AnalyzerResult(false);

            List<AssemblyNameReference> listReferences = new List<AssemblyNameReference>();
            XDocument document = new XDocument();

            XElement rootElement = new XElement("Document");
            document.Add(rootElement);
            XElement assemblyElement = new XElement("Assembly", new XElement("Classes", ""), new XAttribute("Name", assemblyDefinition.Name));
            rootElement.Add(assemblyElement);

            foreach (ModuleDefinition moduleDefinition in assemblyDefinition.Modules)
            {
                ListReferences(listReferences, moduleDefinition);

                foreach (TypeDefinition typeDefinition in moduleDefinition.Types)
                {
                    if (typeDefinition.IsClass)
                    {
                        XElement newElement = new XElement("Class",
                                                new XAttribute("Name", typeDefinition.Name),
                                                
                                                new XAttribute("IsPublic", typeDefinition.IsPublic.ToString()),
                                                new XElement("Fields", ""),
                                                new XElement("Properties", ""),
                                                new XElement("Methods", "")
                                            );

                        bool typeIncludesNetOfficeCalls = AnalyzeEntity(typeDefinition, newElement);
                        if (typeIncludesNetOfficeCalls)
                            assemblyElement.Element("Classes").Add(newElement);                      
                    }
                }
            }


            return new AnalyzerResult(document);
        }

        #endregion

        #region Internal Methods

        private static bool AnalyzeEntity(TypeDefinition typeDefinition, XElement newElement)
        {
            bool result = false;

            foreach (FieldDefinition fieldDefinition in typeDefinition.Fields)
            {
                bool fieldResult = AnalyzeField(fieldDefinition, newElement);
                if (fieldResult)
                    result = fieldResult;
            }

            foreach (MethodDefinition methodDefinition in typeDefinition.Methods)
            {
                bool fieldResult = AnalyzeMethod(methodDefinition, newElement);
                if (fieldResult)
                    result = fieldResult;
            }

            return result;
        }

        private static bool AnalyzeField(FieldDefinition fieldDefinition, XElement newElement)
        {
            bool result = false;
            string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(fieldDefinition.FieldType.FullName);
            if (fieldDefinition.FieldType.FullName.StartsWith(_apiName, StringComparison.InvariantCultureIgnoreCase) && (null != supportByLibrary))
            {
                result = true;
                string typeName = NetOfficeSupportTable.GetName(fieldDefinition.FieldType.FullName);
                string componentName = NetOfficeSupportTable.GetLibrary(fieldDefinition.FieldType.FullName);
                XElement fields = newElement.Element("Fields");
                XElement newField = new XElement(new XElement("Entity", 
                                                 new XAttribute("Type", fieldDefinition.FieldType.FullName),
                                                 new XAttribute("Name", fieldDefinition.Name),
                                                 new XAttribute("IsPublic", fieldDefinition.IsPublic.ToString()),
                                                 new XAttribute("Static", fieldDefinition.IsStatic)));

                XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                newField.Add(supportByNode);
                if (null != supportByLibrary)
                { 
                    fields.Add(newField);
                    supportByNode.Add(new XAttribute("Name", typeName));
                    foreach (string item in supportByLibrary)
                        supportByNode.Add(new XElement("Version", item));
                }
            }
            
            return result;
        }

        private static bool AnalyzeMethod(MethodDefinition methodDefinition, XElement entity)
        {   
            bool result = false;

            bool isProperty = methodDefinition.IsGetter || methodDefinition.IsSetter;

            XElement newMethodNode = new XElement("Method", new XAttribute("Name", methodDefinition.Name), new XAttribute("IsProperty", isProperty.ToString()), new XAttribute("IsPublic", methodDefinition.IsPublic.ToString()), new XElement("FieldSets"), new XElement("LocalFieldSets"));
            XElement parametersNode = null;
            
            // parameter
            foreach (ParameterDefinition paramDefintion in methodDefinition.Parameters)
            {
                string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(paramDefintion.ParameterType.FullName);
                if (paramDefintion.ParameterType.FullName.StartsWith(_apiName) && (null != supportByLibrary))
                {
                    if (null == parametersNode)
                    {
                        parametersNode = new XElement("Parameters");
                        newMethodNode.Add(parametersNode);
                    }
                    string componentName = NetOfficeSupportTable.GetLibrary(paramDefintion.ParameterType.FullName);
                    string typeName = NetOfficeSupportTable.GetName(paramDefintion.ParameterType.FullName);
                    result = true;
                    XElement newParam  = new XElement("Entity",
                                                    new XAttribute("Type", paramDefintion.ParameterType.FullName),
                                                    new XAttribute("Name", paramDefintion.Name),
                                                    new XAttribute("Api", componentName));

                    parametersNode.Add(newParam);

                    XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                    newParam.Add(supportByNode);
                    if (null != supportByLibrary)
                    { 
                        supportByNode.Add(new XAttribute("Name", typeName));
                        foreach (string item in supportByLibrary)
                            supportByNode.Add(new XElement("Version", item));
                    }

                }
            }

            // returnvalue
            if (methodDefinition.ReturnType.FullName.StartsWith(_apiName))
            {
                result = true;
                string componentName = NetOfficeSupportTable.GetLibrary(methodDefinition.ReturnType.FullName);
                string typeName = NetOfficeSupportTable.GetName(methodDefinition.ReturnType.FullName);
                XElement returnValueNode = new XElement("ReturnValue");
                XElement returnValue = new XElement("Entity", new XAttribute("Type", typeName), new XAttribute("FullType", methodDefinition.ReturnType.FullName),
                                                              new XAttribute("Api", componentName));

                XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                returnValue.Add(supportByNode);
                string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(methodDefinition.ReturnType.FullName);
                if (null != supportByLibrary)
                { 
                    supportByNode.Add(new XAttribute("Name", typeName));
                    foreach (string item in supportByLibrary)
                        supportByNode.Add(new XElement("Version", item));
                }

                returnValueNode.Add(returnValue);
                newMethodNode.Add(returnValueNode);
            }

            bool resultVariables = AnalyzeMethodVariables(methodDefinition, entity, newMethodNode);
            bool resultNewObjects = AnalyzeMethodCreateNewObjects(methodDefinition, entity, newMethodNode);
            bool resultCalls = AnalyzeMethodCalls(methodDefinition, entity, newMethodNode);
            bool fieldSetCalls = AnalyzeMethodFieldSets(methodDefinition, entity, newMethodNode);
            bool fieldSetLocalCalls = AnalyzeMethodLocalFieldSets(methodDefinition, entity, newMethodNode);

            if ((result) || (resultVariables) || (resultNewObjects) || (resultCalls) || (fieldSetCalls) || (fieldSetCalls))
            { 
                entity.Element("Methods").Add(newMethodNode);
                result = true;
            }

            return result;
        }
        
        private static bool AnalyzeMethodLocalFieldSets(MethodDefinition methodDefinition, XElement entity, XElement newMethodNode)
        {
            bool result = false;
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            {
                foreach (Instruction itemInstruction in body.Instructions)
                {
                    if ((itemInstruction.OpCode.Name.StartsWith("stloc")) && (!itemInstruction.OpCode.Name.StartsWith("stloc.s")) )
                    {
                        Mono.Cecil.Cil.Instruction methodInstruction = itemInstruction as Mono.Cecil.Cil.Instruction;
                        VariableDefinition variableDefinition = GetLocalVariableDefinition(methodDefinition, itemInstruction);
                        if( (null != variableDefinition) &&  (variableDefinition.VariableType.IsValueType))
                        {
                            Mono.Cecil.Cil.Instruction paramInstruction = GetParameterInstructionForField(methodInstruction);
                            if (null != paramInstruction)
                            {
                                bool sucseed = false;
                                int opValue = GetOperatorValue(paramInstruction, out sucseed);
                                if (sucseed)
                                {
                                    string[] supportByLibrary = _netOfficeSupportTable.GetEnumMemberSupport(variableDefinition.VariableType.FullName, opValue);
                                    if (null != supportByLibrary)
                                    {
                                        XElement newParameter = new XElement("Field", new XAttribute("Name", variableDefinition.ToString()));
                                        string componentName = NetOfficeSupportTable.GetLibrary(variableDefinition.VariableType.FullName);
                                        XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                                        string memberName = _netOfficeSupportTable.GetEnumMemberNameFromValue(variableDefinition.VariableType.FullName, opValue);
                                        supportByNode.Add(new XAttribute("Name", variableDefinition.VariableType.FullName + "." + memberName));
                                        foreach (string item in supportByLibrary)
                                            supportByNode.Add(new XElement("Version", item));
                                        newParameter.Add(supportByNode);
                                        newMethodNode.Element("LocalFieldSets").Add(newParameter);
                                    }
                                }
                            }

                        }
                    }
                }
            }

            return result;
        }

        private static bool AnalyzeMethodFieldSets(MethodDefinition methodDefinition, XElement entity, XElement newMethodNode)
        {
            bool result = false;
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            {
                foreach (Instruction itemInstruction in body.Instructions)
                {
                    if (itemInstruction.OpCode.Name.StartsWith("stfld") || itemInstruction.OpCode.Name.StartsWith("stsfld"))
                    {
                        Mono.Cecil.Cil.Instruction methodInstruction = itemInstruction as Mono.Cecil.Cil.Instruction;
                        Mono.Cecil.FieldDefinition fieldDefinition = methodInstruction.Operand as Mono.Cecil.FieldDefinition;
                        if (fieldDefinition.FieldType.IsValueType)
                        {
                            Mono.Cecil.Cil.Instruction paramInstruction = GetParameterInstructionForField(methodInstruction);
                            if (null != paramInstruction)
                            {
                                bool sucseed = false;
                                int opValue = GetOperatorValue(paramInstruction, out sucseed);
                                if (sucseed)
                                {
                                    string[] supportByLibrary = _netOfficeSupportTable.GetEnumMemberSupport(fieldDefinition.FieldType.FullName, (int)opValue);
                                    if (null != supportByLibrary)
                                    {
                                        XElement newParameter = new XElement("Field", new XAttribute("Name", fieldDefinition.Name));
                                        string componentName = NetOfficeSupportTable.GetLibrary(fieldDefinition.FieldType.FullName);
                                        XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                                        string memberName = _netOfficeSupportTable.GetEnumMemberNameFromValue(fieldDefinition.FieldType.FullName, opValue);
                                        supportByNode.Add(new XAttribute("Name", fieldDefinition.FieldType + "." + memberName));
                                        foreach (string item in supportByLibrary)
                                            supportByNode.Add(new XElement("Version", item));
                                        newParameter.Add(supportByNode);
                                        newMethodNode.Element("FieldSets").Add(newParameter);
                                    }
                                }
                            }

                        }
                    }
                }
            }

            return result;
        }

        private static VariableDefinition GetLocalVariableDefinition(MethodDefinition methodDefinition, Instruction itemInstruction)
        {
            int varIndex = -1;
            switch (itemInstruction.OpCode.ToString())
            {
                case "stloc.0":
                    varIndex = 0;
                    break;
                case "stloc.1":
                    varIndex = 1;
                    break;
                case "stloc.2":
                    varIndex = 2;
                    break;
                case "stloc.3":
                    varIndex = 3;
                    break;
                case "stloc.4":
                    varIndex = 4;
                    break;
                case "stloc.5":
                    varIndex = 5;
                    break;
                case "stloc.6":
                    varIndex = 6;
                    break;
                case "stloc.7":
                    varIndex = 7;
                    break;
                default:
                    varIndex = (int)itemInstruction.Operand;
                    break;
            }

            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            VariableDefinition definiton = body.Variables[varIndex];
            if (definiton.VariableType.FullName.StartsWith(_apiName))
                return definiton;
            else 
                return null;
        }

        private static bool AnalyzeMethodVariables(MethodDefinition methodDefinition, XElement entity, XElement newMethodNode)
        {
            XElement variablesNode = null;

            bool result = false;
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            {
                 // local variables
                 foreach (VariableDefinition itemVariable in body.Variables)
                 {
                     if (itemVariable.VariableType.FullName.StartsWith(_apiName))
                     {
                         string componentName = NetOfficeSupportTable.GetLibrary(itemVariable.VariableType.FullName);
                         string typeName = NetOfficeSupportTable.GetName(itemVariable.VariableType.FullName);

                         if (null == variablesNode)
                         {
                             variablesNode = new XElement("Variables");
                             newMethodNode.Add(variablesNode);
                         }

                         result = true;

                         XElement newVar = new XElement("Entity",
                                                 new XAttribute("Type", itemVariable.VariableType.FullName),
                                                 new XAttribute("Name", itemVariable.ToString()),
                                                 new XAttribute("Api", componentName));

                         variablesNode.Add(newVar);

                         XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                         newVar.Add(supportByNode);
                         string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(itemVariable.VariableType.FullName);
                         if (null != supportByLibrary)
                         {
                             supportByNode.Add(new XAttribute("Name", typeName));
                             foreach (string item in supportByLibrary)
                                 supportByNode.Add(new XElement("Version", item));
                         }
                     }
                 }
            }
             
            return result;
        }

        private static bool AnalyzeMethodCreateNewObjects(MethodDefinition methodDefinition, XElement entity, XElement newMethodNode)
        {
            XElement createNode = null;

            bool result = false;
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            { 
                 // method calls
                foreach (Instruction itemInstruction in body.Instructions)
                {
                    if (itemInstruction.OpCode.Name.StartsWith("newobj"))
                    {
                        Mono.Cecil.MethodReference methodReference = itemInstruction.Operand as Mono.Cecil.MethodReference;
                        string typeName = GetNameFromNewObjMethodReference(methodReference);
                        string componentName = NetOfficeSupportTable.GetLibrary(typeName);
                        string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(typeName);
                        if (typeName.StartsWith(_apiName) && (null != supportByLibrary) )
                        {
                            result = true;
                            if (null == createNode)
                            {
                                createNode = new XElement("NewObjects");
                                newMethodNode.Add(createNode);
                            }

                            XElement newObject = new XElement("Entity",
                                                  new XAttribute("Type", typeName),
                                                  new XAttribute("Api", componentName));

                            createNode.Add(newObject);
                            XElement supportByNode = new XElement("SupportByLibrary");
                            newObject.Add(supportByNode);
                            
                            if (null != supportByLibrary)
                            {
                                supportByNode.Add(new XAttribute("Name", typeName), new XAttribute("Api", componentName));
                                foreach (string item in supportByLibrary)
                                    supportByNode.Add(new XElement("Version", item));
                            }
                        }
                    }
                }
            }

            return result;
        }
         
        private static string[] GetParameter(ParameterDefinition paramDefintion)
        {
            string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(paramDefintion.ParameterType.FullName);
            if (paramDefintion.ParameterType.FullName.StartsWith(_apiName) && (null != supportByLibrary))
                return supportByLibrary;
            else
                return null;
        }

        private static int GetOperatorValue(Mono.Cecil.Cil.Instruction parameterInstruction, out bool sucseed)
        {
            sucseed = true;
            switch (parameterInstruction.OpCode.ToString())
            {
                case "ldc.i4.0":
                    return 0;
                case "ldc.i4.1":
                    return 1;
                case "ldc.i4.2":
                    return 2;
                case "ldc.i4.3":
                    return 3;
                case "ldc.i4.4":
                    return 4;
                case "ldc.i4.5":
                    return 5;
                case "ldc.i4.6":
                    return 7;
                case "ldc.i4.7":
                    return 8;
                default:
                    try
                    {
                        return Convert.ToInt32(parameterInstruction.Operand);     
                    }
                    catch
                    {
                        sucseed = false;
                        return -1;
                    }                                 
            }
        }

        private static int GetOperatorValue(Mono.Cecil.Cil.Instruction parameterInstruction)
        {
            switch (parameterInstruction.OpCode.ToString())
            {
                case "ldc.i4.0":
                    return 0;
                case "ldc.i4.1":
                    return 1;
                case "ldc.i4.2":
                    return 2;
                case "ldc.i4.3":
                    return 3;
                case "ldc.i4.4":
                    return 4;
                case "ldc.i4.5":
                    return 5;
                case "ldc.i4.6":
                    return 7;
                case "ldc.i4.7":
                    return 8;
                default:
                    return Convert.ToInt32(parameterInstruction.Operand);
            }
        }

        private static bool AnalyzeMethodCalls(MethodDefinition methodDefinition, XElement entity, XElement newMethodNode)
        {
            XElement callsNode = null;

            bool result = false;
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            {
                // method calls
                foreach (Instruction itemInstruction in body.Instructions)
                {
                    if (itemInstruction.OpCode.Name.StartsWith("callvirt"))
                    {
                        Mono.Cecil.Cil.Instruction methodInstruction = itemInstruction as Mono.Cecil.Cil.Instruction;
                        Mono.Cecil.MethodReference methodReference = methodInstruction.Operand as Mono.Cecil.MethodReference;
                        string callName = GetCallNameFromAnalyzeMethodCalls(methodReference);
                        string typeName = GetNameFromNewObjMethodReference(methodReference);
                        string componentName = NetOfficeSupportTable.GetLibrary(typeName);
                        
                        if (typeName.StartsWith(_apiName))
                        {
                            string[] supportByLibrary = _netOfficeSupportTable.GetTypeCallSupport(callName);
                            result = true;
                            if (null == callsNode)
                            {
                                callsNode = new XElement("Calls");
                                newMethodNode.Add(callsNode);
                            }

                            XElement newObject = new XElement("Entity",
                                                 new XElement("Parameters"),
                                                 new XAttribute("Type", typeName),
                                                 new XAttribute("Name", callName),
                                                 new XAttribute("Api", componentName));

                            callsNode.Add(newObject);
                            XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                            newObject.Add(supportByNode);
                            if (null != supportByLibrary)
                            {
                                supportByNode.Add(new XAttribute("Name", callName));
                                foreach (string item in supportByLibrary)
                                    supportByNode.Add(new XElement("Version", item));
                            }

                            bool resultParameter = AnalyzeMethodCallParameters(itemInstruction, newObject);
                        }

                        
                    }
                }
            }

            return result;
        }

        private static bool AnalyzeMethodCallParameters(Instruction itemInstruction, XElement newMethodCallNode)
        {
            bool result = false;

            Mono.Cecil.Cil.Instruction methodInstruction = itemInstruction as Mono.Cecil.Cil.Instruction;
            Mono.Cecil.MethodReference methodReference = methodInstruction.Operand as Mono.Cecil.MethodReference;

            int i=1;
            foreach (ParameterDefinition itemParameter in methodReference.Parameters)
            {
                string paramType = itemParameter.ParameterType.FullName;
                if (paramType.StartsWith(_apiName))
                {
                    result = true;
                    Mono.Cecil.Cil.Instruction parameterInstruction = GetParameterInstruction(itemInstruction, i);
                    if (itemParameter.ParameterType.IsValueType)
                    {

                        string[] supportByLibrary = _netOfficeSupportTable.GetEnumMemberSupport(itemParameter.ParameterType.FullName, GetOperatorValue(parameterInstruction));
                        if (null != supportByLibrary)
                        {
                            XElement newParameter = new XElement("Parameter");
                            string componentName = NetOfficeSupportTable.GetLibrary(itemParameter.ParameterType.FullName);
                            XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                            supportByNode.Add(new XAttribute("Name", paramType + "." + GetOperatorValue(parameterInstruction).ToString()));
                            foreach (string item in supportByLibrary)
                                supportByNode.Add(new XElement("Version", item));
                            newParameter.Add(supportByNode);
                            newMethodCallNode.Element("Parameters").Add(newParameter);
                        }

                    }
                    else
                    {
                        string[] supportByLibrary = _netOfficeSupportTable.GetTypeSupport(itemParameter.ParameterType.FullName);
                        if (null != supportByLibrary)
                        {
                            XElement newParameter = new XElement("Parameter");
                            string componentName = NetOfficeSupportTable.GetLibrary(itemParameter.ParameterType.FullName);
                            XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                            supportByNode.Add(new XAttribute("Name", paramType));
                            foreach (string item in supportByLibrary)
                                supportByNode.Add(new XElement("Version", item));
                            newParameter.Add(supportByNode);
                            newMethodCallNode.Element("Parameters").Add(newParameter);
                        }
                    }
                }
                i++;
            }

            return result;
        }
        
        private static Mono.Cecil.Cil.Instruction GetParameterInstructionForField(Instruction itemInstruction)
        {
            Instruction startItem = itemInstruction.Previous;
            while (startItem != null)
            {
                if (startItem.OpCode.Name.StartsWith("ld"))
                {
                    if (startItem.OpCode.Name.StartsWith("ldarg."))
                        return null; 
                    else
                        return startItem;
                }
                startItem = startItem.Previous;
            }
            return null;
        }

        private static Mono.Cecil.Cil.Instruction GetParameterInstruction(Instruction itemInstruction, int index)
        {
            Mono.Cecil.Cil.Instruction methodInstruction = itemInstruction as Mono.Cecil.Cil.Instruction;
            Mono.Cecil.MethodReference methodReference = methodInstruction.Operand as Mono.Cecil.MethodReference;

            int i = 0;
            Instruction startItem = itemInstruction.Previous;
            while (startItem != null)
            {
                if (startItem.OpCode.Name.StartsWith("ld"))
                {
                    i++;
                    if (i == methodReference.Parameters.Count)
                        break;
                }
                startItem = startItem.Previous;
            }

            i = 0;
            while (startItem != null)
            {
                if (startItem.OpCode.Name.StartsWith("ld"))
                {
                    i++;
                    if (i == index)
                        return startItem;
                }
                startItem = startItem.Next;
            }

            return null;
        }
         
        private static VariableDefinition GetVariable(Mono.Cecil.Cil.MethodBody body, Instruction item)
        {
            if (null != item.Operand)
            {
                VariableDefinition returnVar = item.Operand as VariableDefinition;
                return returnVar; 
            }
            else
            { 
                string[] opCodeArray = item.OpCode.Name.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                int varIndex = Convert.ToInt32(opCodeArray[opCodeArray.Length - 1]);
                return body.Variables[varIndex];
            }
        }

        private static bool CheckNetOfficeReferencesExists(AssemblyDefinition assemblyDefinition)
        {
            foreach (ModuleDefinition moduleDefinition in assemblyDefinition.Modules)
            {
                foreach (AssemblyNameReference item in moduleDefinition.AssemblyReferences)
                {
                    switch (item.Name)
                    {
                        case "ExcelApi":
                        case "WordApi":
                        case "OutlookApi":
                        case "PowerPointApi":
                        case "AccessApi":
                        case "OfficeApi":
                            return true;
                    }
                }
            }

            return false;
        }

        private static void ListReferences(List<AssemblyNameReference> listReferences, ModuleDefinition moduleDefinition)
        {
            foreach (AssemblyNameReference item in moduleDefinition.AssemblyReferences)
            {
                switch (item.Name)
                {
                    case "ExcelApi":
                    case "WordApi":
                    case "OutlookApi":
                    case "PowerPointApi":
                    case "AccessApi":
                    case "OfficeApi":
                        if (!listReferences.Contains(item))
                            listReferences.Add(item);
                        break;
                }
            }
        }

        private static string GetCallNameFromAnalyzeMethodCalls(Mono.Cecil.MethodReference methodReference)
        {
            string[] array = methodReference.ToString().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            return array[1].Trim();
        }

        private static string GetNameFromNewObjMethodReference(Mono.Cecil.MethodReference methodReference)
        {
            string[] array = methodReference.ToString().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
            string[] array2 = array[1].Split(new string[] { "::" }, StringSplitOptions.RemoveEmptyEntries);
            return array2[0].Trim();
        }

        #endregion

    }
}
