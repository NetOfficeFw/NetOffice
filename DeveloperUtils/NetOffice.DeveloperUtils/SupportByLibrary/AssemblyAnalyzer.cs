using System;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using Mono.Cecil;
using Mono.Cecil.Cil;

namespace NetOffice.DeveloperUtils.SupportByLibrary
{
    public class AssemblyAnalyzerSettingsLibrary
    {
        public bool Version9 { get; set; }
        public bool Version10 { get; set; }
        public bool Version11 { get; set; }
        public bool Version12 { get; set; }
        public bool Version14 { get; set; }
    }

    public class AssemblyAnalyzerSettings
    {
        AssemblyAnalyzerSettingsLibrary _excel;
        AssemblyAnalyzerSettingsLibrary _word;
        AssemblyAnalyzerSettingsLibrary _outlook;
        AssemblyAnalyzerSettingsLibrary _powerpoint;
        AssemblyAnalyzerSettingsLibrary _access;
        AssemblyAnalyzerSettingsLibrary _office;

        public AssemblyAnalyzerSettings()
        {
            _excel = new AssemblyAnalyzerSettingsLibrary();
            _word = new AssemblyAnalyzerSettingsLibrary();
            _outlook = new AssemblyAnalyzerSettingsLibrary();
            _powerpoint = new AssemblyAnalyzerSettingsLibrary();
            _access = new AssemblyAnalyzerSettingsLibrary();
            _office = new AssemblyAnalyzerSettingsLibrary();
        }

        public AssemblyAnalyzerSettingsLibrary Excel
        {
            get { return _excel; }
        }

        public AssemblyAnalyzerSettingsLibrary Word
        {
            get { return _word; }
        }

        public AssemblyAnalyzerSettingsLibrary Outlook
        {
            get { return _outlook; }
        }

        public AssemblyAnalyzerSettingsLibrary PowerPoint
        {
            get { return _powerpoint; }
        }

        public AssemblyAnalyzerSettingsLibrary Access
        {
            get { return _access; }
        }

        public AssemblyAnalyzerSettingsLibrary Office
        {
            get { return _access; }
        }
    }

    public static class AssemblyAnalyzer
    {
        private static string _apiName = "NetOffice";

        #region Public Methods

        /// <summary>
        /// Analzye an assembly for NetOffice calls
        /// </summary>
        /// <param name="fullFileName"></param>
        /// <param name="analyzeDependencies"></param>
        public static string[] AnalyzeAssembly(AssemblyDefinition assemblyDefinition, AssemblyAnalyzerSettings settings)
        {
            List<AssemblyNameReference> listReferences = new List<AssemblyNameReference>();

            XDocument document = new XDocument();
            XElement rootElement = new XElement("Assembly", new XElement("Classes", ""), new XAttribute("Name", assemblyDefinition.Name));
            document.Add(rootElement);

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
                        case "VBIDEApi":
                        case "DAOApi":
                        case "ADODBApi":
                        case "OWC10Api":
                        case "MSDATASRCApi":
                        case "MSComctlLibApi":
                            if(!listReferences.Contains(item))
                                listReferences.Add(item);
                            break;
                    }
                }
           
                foreach (TypeDefinition typeDefinition in moduleDefinition.Types)
                {
                    if (typeDefinition.IsClass)
                    {
                        XElement newElement = new XElement("Class",
                                                new XAttribute("Name", typeDefinition.Name),
                                                new XElement("Fields", ""),
                                                new XElement("Properties", ""),
                                                new XElement("Methods", "")
                                            );

                        bool typeIncludesNetOfficeCalls = AnalyzeEntity(typeDefinition, newElement);
                        if (typeIncludesNetOfficeCalls)
                            rootElement.Element("Classes").Add(newElement);
                    }
                }
            }

            string result = NetOfficeAnalyzer.AnalyzeNetOfficeAssemblies(document, listReferences, settings);
            return new string[]{result, document.ToString()};
        }

        #endregion

        #region Internal Methods
        
        private static bool AnalyzeField(FieldDefinition fieldDefinition, XElement newElement)
        {
            bool result = false;
            if (fieldDefinition.FieldType.FullName.StartsWith(_apiName, StringComparison.InvariantCultureIgnoreCase))
            {
                result = true;

                newElement.Element("Fields").Add(new XElement("Field",
                                                    new XAttribute("Type", fieldDefinition.FieldType.FullName),
                                                    new XAttribute("Name", fieldDefinition.Name)));
            }
            return result;
        }

        private static bool AnalyzeMethod(MethodDefinition methodDefinition, XElement entity)
        {
            bool result = false;

            XElement newMethodNode = new XElement("Method", new XAttribute("Name", methodDefinition.Name));

            // parameter
            foreach (ParameterDefinition paramDefintion in methodDefinition.Parameters)
            {
                if (paramDefintion.ParameterType.FullName.StartsWith(_apiName))
                {
                    result = true;
                }

                newMethodNode.Add(new XElement("Parameter",
                                                    new XAttribute("Type", paramDefintion.ParameterType.FullName),
                                                    new XAttribute("Name", paramDefintion.Name)));
            }

            // returnvalue
            if (methodDefinition.ReturnType.FullName.StartsWith(_apiName))
            {
                result = true;
                newMethodNode.Add(new XElement("ReturnValue", new XAttribute("Type", methodDefinition.ReturnType.FullName)));
            }

            // analyze body
            Mono.Cecil.Cil.MethodBody body = methodDefinition.Body;
            if (null != body)
            {
                // local variables
                foreach (VariableDefinition item in body.Variables)
                {
                    if (item.VariableType.FullName.StartsWith(_apiName))
                    {
                        result = true;

                        XElement newVar = new XElement("Var",
                                                new XAttribute("Type", item.VariableType.FullName),
                                                new XAttribute("Name", item.ToString()));

                        newMethodNode.Add(newVar);
                    }
                }

                // method calls
                foreach (Instruction item in body.Instructions)
                {
                    if (item.OpCode.Name.StartsWith("stloc"))
                    {
                        bool setResult = Procced_SetToLocalVariable(newMethodNode, body, item);
                        if (setResult)
                            result = true;
                    }
                    else if (item.OpCode.Name.StartsWith("stfld"))
                    {
                        bool setResult = Procced_SetToField(newMethodNode, body, item);
                        if (setResult)
                            result = true;
                    }
                    else if (item.Operand is Mono.Cecil.MethodReference)
                    {
                        Mono.Cecil.MethodReference methodReference = item.Operand as Mono.Cecil.MethodReference;
                        Mono.Cecil.TypeReference typeReference = methodReference.ReturnType;

                        if (methodReference.DeclaringType.FullName.StartsWith(_apiName))
                        {
                            result = true;

                            XElement newMethod = new XElement("Call",
                                                    new XAttribute("Type", methodReference.DeclaringType.FullName),
                                                    new XAttribute("Name", methodReference.Name)
                                                    );

                            newMethodNode.Add(newMethod);

                            Procced_MethodParams(newMethod, body, item, methodReference);
                        }
                        else
                        {
                            XElement newMethod = new XElement("Call",
                                                       new XAttribute("Type", methodReference.DeclaringType.FullName),
                                                       new XAttribute("Name", methodReference.Name)
                                                       );
                            bool paramResult = Procced_MethodParams(newMethod, body, item, methodReference);
                            if (paramResult)
                            {
                                newMethodNode.Add(newMethod);
                                result = true;
                            }

                        }
                    }

                }
            }

            if (result)
                entity.Element("Methods").Add(newMethodNode);

            return result;
        }

        private static bool AnalyzeProperty(PropertyDefinition propertyDefinition, XElement entity)
        {
            bool result = false;

            XElement newPropertyNode = new XElement("Property",
                                                   new XAttribute("Type", propertyDefinition.PropertyType.FullName),
                                                   new XAttribute("Name", propertyDefinition.Name));

            // parameter
            foreach (ParameterDefinition paramDefintion in propertyDefinition.GetMethod.Parameters)
            {
                if (paramDefintion.ParameterType.FullName.StartsWith("NetOffice"))
                {
                    result = true;

                    newPropertyNode.Add(new XElement("Parameter",
                         new XAttribute("Type", paramDefintion.ParameterType.FullName),
                         new XAttribute("Name", paramDefintion.Name)
                        ));
                }
            }

            // returnvalue
            if (propertyDefinition.GetMethod.ReturnType.FullName.StartsWith("NetOffice"))
            {
                result = true;
                newPropertyNode.Add(new XElement("ReturnValue", new XAttribute("Type", propertyDefinition.GetMethod.ReturnType.FullName)));
            }

            Mono.Cecil.Cil.MethodBody body = propertyDefinition.GetMethod.Body;
            if (null != body)
            {
                // local variables
                foreach (VariableDefinition item in body.Variables)
                {
                    if (item.VariableType.FullName.StartsWith(_apiName))
                    {
                        result = true;

                        XElement newVar = new XElement("Var",
                                                new XAttribute("Type", item.VariableType.FullName),
                                                new XAttribute("Name", item.ToString()));

                        newPropertyNode.Add(newVar);
                    }
                }

                // method calls
                foreach (Instruction item in body.Instructions)
                {
                    if (item.OpCode.Name.StartsWith("stloc"))
                    {
                        bool setResult = Procced_SetToLocalVariable(newPropertyNode, body, item);
                        if (setResult)
                            result = true;
                    }
                    else if (item.OpCode.Name.StartsWith("stfld"))
                    {
                        bool setResult = Procced_SetToField(newPropertyNode, body, item);
                        if (setResult)
                            result = true;
                    }
                    else if (item.Operand is Mono.Cecil.MethodReference)
                    {
                        Mono.Cecil.MethodReference methodReference = item.Operand as Mono.Cecil.MethodReference;
                        Mono.Cecil.TypeReference typeReference = methodReference.ReturnType;

                        if (methodReference.DeclaringType.FullName.StartsWith(_apiName))
                        {
                            result = true;

                            XElement newMethod = new XElement("Call",
                                                    new XAttribute("Type", methodReference.DeclaringType.FullName),
                                                    new XAttribute("Name", methodReference.Name)
                                                    );

                            newPropertyNode.Add(newMethod);

                            Procced_MethodParams(newMethod, body, item, methodReference);
                        }
                        else
                        {
                            XElement newMethod = new XElement("Call",
                                                       new XAttribute("Type", methodReference.DeclaringType.FullName),
                                                       new XAttribute("Name", methodReference.Name)
                                                       );
                            bool paramResult = Procced_MethodParams(newMethod, body, item, methodReference);
                            if (paramResult)
                            {
                                newPropertyNode.Add(newMethod);
                                result = true;
                            }

                        }
                    }

                }
            }

            if (result)
                entity.Element("Properties").Add(newPropertyNode);

            return result;
        }

        private static bool AnalyzeEntity(TypeDefinition typeDefinition, XElement newElement)
        {
            bool result = false;

            foreach (FieldDefinition fieldDefinition in typeDefinition.Fields)
            {
                bool fieldResult = AnalyzeField(fieldDefinition, newElement);
                if (fieldResult)
                    result = fieldResult;
            }

            foreach (PropertyDefinition propertyDefinition in typeDefinition.Properties)
            {
                bool fieldResult = AnalyzeProperty(propertyDefinition, newElement);
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

        private static bool Procced_SetToField(XElement methodNode, Mono.Cecil.Cil.MethodBody body, Instruction item)
        {
            bool result = false;

            Mono.Cecil.FieldDefinition fieldDefinition = item.Operand as Mono.Cecil.FieldDefinition;
            if (fieldDefinition.FieldType.FullName.StartsWith(_apiName))
            {
                Instruction prevItem = item.Previous;
                int opCodeValue = 0;
                bool valueIsSet = false;
                if (prevItem.OpCode.Name == "ldc.i4.s")
                {
                    // more than 7 value
                    opCodeValue = Convert.ToInt32(prevItem.Operand);
                    valueIsSet = true;
                }
                else if (prevItem.OpCode.Name.StartsWith("ldc.i4"))
                {
                    string[] opCodeArray = prevItem.OpCode.Name.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    if (opCodeArray.Length == 2)
                    {
                        opCodeValue = Convert.ToInt32(prevItem.Operand);
                        valueIsSet = true;
                    }
                    else
                    {
                        opCodeValue = Convert.ToInt32(opCodeArray[opCodeArray.Length - 1]);
                        valueIsSet = true;
                    }
                }

                if (true == valueIsSet)
                {
                    result = true;
                    methodNode.Add(new XElement("FieldSet", new XAttribute(fieldDefinition.Name, opCodeValue)));
                }
            }

            return result;
        }

        private static bool Procced_SetToLocalVariable(XElement methodNode, Mono.Cecil.Cil.MethodBody body, Instruction item)
        {
            bool result = false;

            // set to local variable
            Instruction prevItem = item.Previous;
            while (!prevItem.OpCode.Name.StartsWith("ldc") && !prevItem.OpCode.Name.StartsWith("ldnull")
                && !prevItem.OpCode.Name.StartsWith("newobj") && !prevItem.OpCode.Name.StartsWith("ldfld")
                && !prevItem.OpCode.Name.StartsWith("call") && !prevItem.OpCode.Name.StartsWith("ldstr")
                && !prevItem.OpCode.Name.StartsWith("ldsfld") && !prevItem.OpCode.Name.StartsWith("ldarg"))

                prevItem = prevItem.Previous;

            VariableDefinition localVariable = GetVariable(body, item);
            if (localVariable.VariableType.FullName.StartsWith(_apiName))
            {
                int opCodeValue = 0;
                bool valueIsSet = false;
                if (prevItem.OpCode.Name == "ldc.i4.s")
                {
                    // more than 7 value
                    opCodeValue = Convert.ToInt32(prevItem.Operand);
                    valueIsSet = true;
                }
                else if (prevItem.OpCode.Name.StartsWith("ldc.i4"))
                {
                    string[] opCodeArray = prevItem.OpCode.Name.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    if (opCodeArray.Length == 2)
                    {
                        opCodeValue = Convert.ToInt32(prevItem.Operand);
                        valueIsSet = true;
                    }
                    else
                    {
                        opCodeValue = Convert.ToInt32(opCodeArray[opCodeArray.Length - 1]);
                        valueIsSet = true;
                    }
                }

                if (true == valueIsSet)
                {
                    result = true;
                    XElement varNode = (from a in methodNode.Elements("Var")
                                        where a.Attribute("Name").Value.Equals(localVariable.ToString())
                                        select a).FirstOrDefault();

                    varNode.Add(new XElement("Set", new XAttribute("Value", opCodeValue)));
                }
            }

            return result;
        }

        private static bool Procced_MethodParams(XElement newMethod, Mono.Cecil.Cil.MethodBody body, Instruction item, Mono.Cecil.MethodReference methodReference)
        {
            bool result = false;
            int i = 0;
            Instruction prevItem = item;
            while (i < methodReference.Parameters.Count)
            {
                if (methodReference.Parameters[i].ParameterType.FullName.StartsWith(_apiName))
                    result = true;

                prevItem = prevItem.Previous;
                if(null == prevItem)
                    break;

                if (prevItem.OpCode.Name.StartsWith("ldloc"))
                {
                    VariableDefinition localVariable = GetVariable(body, prevItem);
                    if (true == localVariable.VariableType.IsValueType)
                    {
                        newMethod.Add(new XElement("Param", new XAttribute("Var", localVariable.Name)));
                    }
                    else
                    {
                        newMethod.Add(new XElement("Param", new XAttribute("Object", localVariable.VariableType.FullName)));
                    }
                    i++;
                }
                else if (prevItem.OpCode.Name.StartsWith("ldarg"))
                {
                    if ("ldfld" == prevItem.Next.OpCode.Name)
                    {
                        FieldDefinition field = prevItem.Next.Operand as FieldDefinition;
                        if (true == field.FieldType.IsValueType)
                        {
                            newMethod.Add(new XElement("Param", new XAttribute("Field", field.Name)));
                        }
                        else
                        {
                            newMethod.Add(new XElement("Param", new XAttribute("Object", field.FieldType.FullName)));
                        }

                        i++;
                    }
                }
                else if (prevItem.OpCode.Name.StartsWith("newobj"))
                {
                    MethodReference refd = prevItem.Operand as MethodReference;
                    newMethod.Add(new XElement("Param", new XAttribute("Object", refd.DeclaringType.FullName)));
                    i++;
                }
                else if (prevItem.OpCode.Name == "ldc.i4.s")
                {
                    int opCodeValue = Convert.ToInt32(prevItem.Operand);
                    newMethod.Add(new XElement("Param", new XAttribute("Value", opCodeValue)));
                    i++;
                }
                else if (prevItem.OpCode.Name.StartsWith("ldc.i4"))
                {
                    int opCodeValue = 0;
                    string[] opCodeArray = prevItem.OpCode.Name.Split(new string[] { "." }, StringSplitOptions.RemoveEmptyEntries);
                    if (opCodeArray.Length == 2)
                    {
                        opCodeValue = Convert.ToInt32(prevItem.Operand);
                    }
                    else
                    {
                        if ("m1" == opCodeArray[opCodeArray.Length - 1])
                            opCodeValue = -1;
                        else
                            opCodeValue = Convert.ToInt32(opCodeArray[opCodeArray.Length - 1]);
                    }

                    newMethod.Add(new XElement("Param", new XAttribute("Value", opCodeValue)));
                    i++;
                }
                else if (prevItem.OpCode.Name == "ldnull")
                {
                    newMethod.Add(new XElement("Param", new XAttribute("Value", "null")));
                    i++;
                }
            }

            return result;
        }

        #endregion
    }
}
