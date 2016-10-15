using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.Runtime.InteropServices;

namespace RegAddin.Metrics
{
    internal class AddinMetrics
    {
        private enum PossibleAddin
        {
            No = 0,
            NetOffice = 1,
            Interop = 2
        }

        private static Guid _extensibilityID = new Guid("b65ad801-abaf-11d0-bb8b-00a0c90f2744");
        private static string _netofficeExtensibilityName = "NetOffice.Tools.IDTExtensibility2";
        private static string _netOfficeRegisterAttributeName = "NetOffice.Tools.RegisterFunctionAttribute";
        private static string _netOfficeUnregisterAttributeName = "NetOffice.Tools.UnRegisterFunctionAttribute";
        private static string _netOfficeRegisterErrorAttributeName = "NetOffice.Tools.RegisterErrorHandler";


        private static Guid _netOfficeRegisterCallID = new Guid("D8FAB9D7-10D1-4AA3-8DBA-D9CCA8C4CE9B");
        private static Guid _netOfficeRegisterScopeID = new Guid("FC5DC88D-D4D8-4BC8-A206-F55E7CD94C89");
        private static Guid _netOfficeRegisterErrorMethodKindID = new Guid("F9A44508-4DC1-4E30-8195-0AFED88288E5");

        internal AddinMetrics(Assembly assembly, bool useWindow)
        {
            Assembly = assembly;
            Result = new Dictionary<string, bool>();
            UseWindow = useWindow;
        }

        private bool UseWindow { get; set; }

        private Assembly Assembly { get; set; }

        private object[] Attributes { get; set; }

        private AssemblyName Name { get; set; }

        private Type[] Types { get; set; }

        private Dictionary<string, bool> Result { get; set; }

        internal void Check()
        {
            Result.Clear();
            Name = Assembly.GetName();
            Types = Assembly.GetTypes();
            Attributes = Assembly.GetCustomAttributes(true);
            AnalyzeAssembly();
            AnalyzeClasses();

            if (Result.ContainsValue(false))
            {
                if (UseWindow)
                    new WindowPresenter().Show(Result);
                else
                    new ConsolePresenter().Show(Result);
            }
        }

        private void AnalyzeAssembly()
        {
            bool comVisible = false;
            bool hasGuid = false;
            bool isSigned = false;
            foreach (object item in Attributes)
            {
                ComVisibleAttribute comAttribute = item as ComVisibleAttribute;
                if (null != comAttribute && comAttribute.Value)
                    comVisible = true;

                GuidAttribute guidAttribute = item as GuidAttribute;
                if (null != guidAttribute)
                    hasGuid = true;
            }

            isSigned = (Name.GetPublicKeyToken().Length != 0);

            Result.Add("Assembly-ComVisible", comVisible);
            Result.Add("Assembly-HasGuid", hasGuid);
            Result.Add("Assembly-IsSigned", isSigned);
        }

        private void AnalyzeClasses()
        {
            foreach (Type type in Types)
            {
                if (!type.IsClass)
                    continue;
                PossibleAddin possible = IsPossibleAddinConnectClass(type);
                if (possible == PossibleAddin.No)
                    continue;

                switch (possible)
                {                   
                    case PossibleAddin.NetOffice:
                        AnalyzeNetOfficeConnectClass(type);
                        break;
                    case PossibleAddin.Interop:
                        AnalyzeInteropConnectClass(type);
                        break;
                    case PossibleAddin.No:
                        break;
                    default:
                        throw new IndexOutOfRangeException();
                }
            }
        }

        private void AnalyzeNetOfficeConnectClass(Type addin)
        {
            if (!addin.IsPublic)
                Result.Add("Addin Module " + addin.Name + " Not Public" , false);

            if(!addin.GetConstructors().Any(e => e.IsPublic && e.GetParameters().Length == 0))
                Result.Add("Addin Module " + addin.Name + " Missing Ctor", false);

            IEnumerable<object> customAttributes = Common.AttributeReflection.GetCustomClassAttributes(addin);

            if(!Common.AttributeReflection.ComVisibleAttributeExists(customAttributes))
                Result.Add("Addin Module " + addin.Name + " Missing ComVisible Attribute", false);
            if (!Common.AttributeReflection.GuidAttributeExists(customAttributes))
                Result.Add("Addin Module " + addin.Name + " Missing Guid Attribute", false);
            if (!Common.AttributeReflection.ProgIdAttributeExists(customAttributes))
                Result.Add("Addin Module " + addin.Name + " Missing ProgId Attribute", false);
            if (!Common.AttributeReflection.ClassInterfaceAttributeExists(customAttributes))
                Result.Add("Addin Module " + addin.Name + " Missing ClassInterface Attribute", false);

            MethodInfo[] methods = addin.GetMethods(BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            MethodInfo regMethod = GetNetOfficeCustomRegMethod(methods);
            MethodInfo unregMethod = GetNetOfficeCustomUnregMethod(methods);
            MethodInfo regErrorMethod = GetNetOfficeRegisterErrorMethod(methods);

            if (null != regMethod)
            {
                if(!regMethod.IsStatic)
                    Result.Add("Addin Module " + addin.Name + " Register Method Is Not Static", false);
                if (!RegisterMethodHasValidSignatur(regMethod))
                    Result.Add("Addin Module " + addin.Name + " Register Method Has Invalid Argument Signature", false);
            }

            if (null != unregMethod)
            {
                if (!unregMethod.IsStatic)
                    Result.Add("Addin Module " + addin.Name + " Unregister Method Is Not Static", false);
                if (!UnregisterMethodHasValidSignatur(regMethod))
                    Result.Add("Addin Module " + addin.Name + " Unregister Method Has Invalid Argument Signature", false);
            }

            if (null != regErrorMethod)
            {
                if (!regErrorMethod.IsStatic)
                    Result.Add("Addin Module " + addin.Name + " Register Erorr Handler Method Is Not Static", false);
                if (!RegisterErrorMethodHasValidSignatur(regMethod))
                    Result.Add("Addin Module " + addin.Name + " Register Erorr Handler Method Has Invalid Argument Signature", false);

            }
        }

        private static bool RegisterErrorMethodHasValidSignatur(MethodInfo item)
        {
            ParameterInfo[] arguments = item.GetParameters();
            if (arguments.Length == 0)
                return true;
            if (arguments.Length == 1)
                return (arguments[0].ParameterType.GUID == _netOfficeRegisterErrorMethodKindID);
            else if (arguments.Length == 2)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.GUID == _netOfficeRegisterErrorMethodKindID)
                        && (arguments[1].ParameterType.Name == "System.Exception");
            }
            else if (arguments.Length == 3)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.GUID == _netOfficeRegisterErrorMethodKindID)
                        && (arguments[1].ParameterType.Name == "System.Exception")
                        && (arguments[2].ParameterType.GUID == _netOfficeRegisterScopeID);
            }
            return false;
        }

        private static bool UnregisterMethodHasValidSignatur(MethodInfo item)
        {
            ParameterInfo[] arguments = item.GetParameters();
            if (arguments.Length == 0)
                return true;
            if (arguments.Length == 1)
                return arguments[0].ParameterType.FullName == "System.Type";
            else if (arguments.Length == 2)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.FullName == "System.Type") 
                        && (arguments[1].ParameterType.GUID == _netOfficeRegisterCallID);
            }
            else if(arguments.Length == 3)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.FullName == "System.Type") 
                        && (arguments[1].ParameterType.GUID == _netOfficeRegisterCallID)
                        && (arguments[2].ParameterType.GUID == _netOfficeRegisterScopeID);
            }
            return false;
        }

        private static bool RegisterMethodHasValidSignatur(MethodInfo item)
        {
            ParameterInfo[] arguments = item.GetParameters();
            if (arguments.Length == 0)
                return true;
            if (arguments.Length == 1)
                return arguments[0].ParameterType.FullName == "System.Type";
            else if (arguments.Length == 2)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.FullName == "System.Type")
                        && (arguments[1].ParameterType.GUID == _netOfficeRegisterCallID);
            }
            else if (arguments.Length == 3)
            {
                if (arguments.Length == 1)
                    return (arguments[0].ParameterType.FullName == "System.Type")
                        && (arguments[1].ParameterType.GUID == _netOfficeRegisterCallID)
                        && (arguments[2].ParameterType.GUID == _netOfficeRegisterScopeID);
            }
            return false;
        }

        private static bool HasAttribute(MethodInfo item, string fullAttributeName)
        {
            object[] attributes = item.GetCustomAttributes(false);
            foreach (var attribute in attributes)
            {
                Attribute attrib = attribute as Attribute;
                Type attribType = attrib.TypeId as Type;
                if (attribType.FullName == fullAttributeName)
                    return true;

            }
            return false;
        }

        private MethodInfo GetNetOfficeRegisterErrorMethod(MethodInfo[] methods)
        {
            foreach (MethodInfo item in methods)
            {
                if (HasAttribute(item, _netOfficeRegisterErrorAttributeName))
                    return item;
            }
            return null;
        }

        private MethodInfo GetNetOfficeCustomRegMethod(MethodInfo[] methods)
        {
            foreach (MethodInfo item in methods)
            {
                if (HasAttribute(item, _netOfficeRegisterAttributeName))
                    return item;
            }
            return null;
        }

        private MethodInfo GetNetOfficeCustomUnregMethod(MethodInfo[] methods)
        {
            foreach (MethodInfo item in methods)
            {
                if (HasAttribute(item, _netOfficeUnregisterAttributeName))
                    return item;
            }
            return null;
        }

        private void AnalyzeInteropConnectClass(Type addin)
        {

        }
         
        private static PossibleAddin IsPossibleAddinConnectClass(Type item)
        {
            Type type = item;
            while (null != type)
            {
                Type[] implInterfaces = type.GetInterfaces();
                foreach (Type interfaceType in implInterfaces)
                {
                    if(interfaceType.GUID == _extensibilityID)
                    {
                        if (interfaceType.FullName == _netofficeExtensibilityName)
                            return PossibleAddin.NetOffice;
                        else
                            return PossibleAddin.Interop;
                    }
                }

                if (null != type.BaseType && type.BaseType.FullName == "System.Object")
                    break;
                type = type.BaseType;
            }

            return PossibleAddin.No;
        }

        /*
         Assembly

         // Assembly ist signiert 
         // Assembly ist COMVisible
         // Assembly hat ID

         Addin Klasse - Interop Assemblies

        // offentlich
        // öffentlicher parameterlose ctor
        // ComVisible, GuidAttribute, ProgIdAttribute, ClassInterfaceAttribute ???        
        // Register/Unregister Methoden - wenn vorhanden - sind public static mit richtiger signatur
        // OnError bzw. OnRegisterError auch noch prüfen

        // Abgeleitete Klasse Implementiert nicht IDTExtensibility2

         Addin Klasse - NetOffice
         // offentlich
         // öffentlicher parameterlose ctor
         // GuidAttribute, ProgIdAttribute, COMAddinAttribute
         // Register/Unregister Methoden - wenn vorhanden - sind static mit richtiger signatur
         // Com Register Attribute abgeleiteter Klasse darf nicht sein
        */
    }
}
