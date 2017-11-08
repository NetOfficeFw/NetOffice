using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Mono.Cecil;
using Mono.Cecil.Cil;

namespace NetOffice.DeveloperToolbox.ToolboxControls.OfficeCompatibility
{
    public static class AssemblyAnalyzer
    {
        #region Fields

        private static string _apiName = "NetOffice";
        private static NetOfficeSupportTable _netOfficeSupportTable;

        #endregion
        
        #region Ctor

        static AssemblyAnalyzer()
        {
            _netOfficeSupportTable = new NetOfficeSupportTable();
        }

        #endregion

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
            if (fieldDefinition.FieldType.FullName.StartsWith(_apiName, StringComparison.InvariantCultureIgnoreCase) &&
                !fieldDefinition.FieldType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase)  && CountOf(fieldDefinition.FieldType.FullName, ".") > 1 && (null != supportByLibrary))
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
                if (paramDefintion.ParameterType.FullName.StartsWith(_apiName) && (!paramDefintion.ParameterType.FullName.StartsWith("NetOffice.DeveloperToolbox")) &&
                    !paramDefintion.ParameterType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase) && CountOf(paramDefintion.ParameterType.FullName, ".") > 1 && (null != supportByLibrary))
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
            if ((methodDefinition.ReturnType.FullName.StartsWith(_apiName) &&
                !methodDefinition.ReturnType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase) &&
                CountOf(methodDefinition.ReturnType.FullName, ".") > 1 && 
                (!methodDefinition.ReturnType.FullName.StartsWith("NetOffice.DeveloperToolbox"))))
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
                        if (fieldDefinition != null && fieldDefinition.FieldType.IsValueType)
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
            if (definiton.VariableType.FullName.StartsWith(_apiName) && CountOf(definiton.VariableType.FullName, ".") > 1 && !definiton.VariableType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase))
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
                     if (itemVariable.VariableType.FullName.StartsWith(_apiName) && CountOf(itemVariable.VariableType.FullName, ".") > 1 && !itemVariable.VariableType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase))
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
                        if (typeName.StartsWith(_apiName) && !typeName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase)  &&
                             CountOf(typeName, ".") > 1 && (null != supportByLibrary))
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
            if (paramDefintion.ParameterType.FullName.StartsWith(_apiName) && !paramDefintion.ParameterType.FullName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase) &&
                CountOf(paramDefintion.ParameterType.FullName, ".") > 1 && (null != supportByLibrary))
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

                        string[] testArray = typeName.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                        if (testArray.Length >= 3 && testArray[2] == "Tools")
                            continue;
                        if ((typeName.StartsWith(_apiName)) && !typeName.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase) &&
                            CountOf(typeName, ".") > 1 && (!typeName.StartsWith("NetOffice.DeveloperToolbox")))
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
                if (paramType.StartsWith(_apiName) && CountOf(paramType, ".") > 1 && !paramType.Equals(_apiName + ".dll", StringComparison.InvariantCultureIgnoreCase))
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

                            string enumMemberName = _netOfficeSupportTable.GetEnumMemberNameFromValue(itemParameter.ParameterType.FullName, GetOperatorValue(parameterInstruction));
                            if (null != enumMemberName)
                                supportByNode.Add(new XAttribute("Name", paramType + "." + enumMemberName));
                            else
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
                else
                {
                    Instruction prevInstruction = methodInstruction.Previous;
                    while (prevInstruction.OpCode.Name.StartsWith("ld") || prevInstruction.OpCode.Name.StartsWith("box"))
                    {
                        if (null != prevInstruction.Operand)
                        { 
                            string targetName = prevInstruction.Operand.ToString();
                            string[] dumyByLibrary = _netOfficeSupportTable.GetTypeSupport(targetName);
                            if (null != dumyByLibrary)
                            {
                                if(null != prevInstruction.Previous.Operand)
                                {  
                                    int temp = 0;
                                    if (Int32.TryParse(prevInstruction.Previous.Operand.ToString(), out temp))
                                    {
                                        int opValue = Convert.ToInt32(temp);
                                        string enumMemberName = _netOfficeSupportTable.GetEnumMemberNameFromValue(targetName, opValue);
                                        if (null != enumMemberName)
                                        {
                                            string[] supportByLibrary = _netOfficeSupportTable.GetEnumMemberSupport(targetName, opValue);
                                            if (null != supportByLibrary)
                                            {
                                                XElement newParameter = new XElement("Parameter");
                                                string componentName = NetOfficeSupportTable.GetLibrary(targetName);
                                                XElement supportByNode = new XElement("SupportByLibrary", new XAttribute("Api", componentName));
                                                supportByNode.Add(new XAttribute("Name", targetName + "." + enumMemberName));

                                                foreach (string item in supportByLibrary)
                                                    supportByNode.Add(new XElement("Version", item));
                                                newParameter.Add(supportByNode);
                                                newMethodCallNode.Element("Parameters").Add(newParameter);
                                            }
                                        }
                                    }
                                } 
                            }
                        }
                        prevInstruction = prevInstruction.Previous;
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
                        case "MSProjectApi":
                        case "VisioApi":
                        case "PublisherApi":
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
                    case "MSProjectApi":
                    case "VisioApi":
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

        private static int CountOf(string value, string targetExpression)
        {
            string[] splitArray = value.Split(new string[] { targetExpression }, StringSplitOptions.RemoveEmptyEntries);
            return splitArray.Length - 1;
        }

        #endregion
    }
}