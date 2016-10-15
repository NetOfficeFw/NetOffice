using System;
using System.Collections.Generic;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Linq;
using System.Text;

namespace RegAddin.Common
{
    internal class AddinClassInformations
    {
        private AddinClassInformations(string assemblyName, string assemblyVersion, string assemblyCulture, string assemblyToken,
            string runtimeVersion, string[] classesRoot, string progId, string fullClassName, string id, string codebase)
        {
            AssemblyName = assemblyName;
            AssemblyVersion = assemblyVersion;
            AssemblyCulture = assemblyCulture;
            AssemblyToken = assemblyToken;
            RuntimeVersion = runtimeVersion;
            ClassesRoot = classesRoot;
            ProgId = progId;
            FullClassName = fullClassName;
            Id = id;
            Codebase = codebase;
            ComponentCategoryId = "62C8FE65-4EBB-45E7-B440-6E39B2CDBF29";
        }

        internal static AddinClassInformations Create(Assembly addinAssembly, IEnumerable<object> assemblyAttributes,
                       SingletonSettings.UnRegisterMode mode, Type addinClassType, IEnumerable<object> addinClassAttributes)
        {
            AssemblyName binaryHeader = addinAssembly.GetName();
            string assemblyName = binaryHeader.Name;
            string assemblyVersion = binaryHeader.Version.ToString();
            string assemblyCulture = CultureInfoConversion.ConvertToString(binaryHeader.CultureInfo);
            string assemblyToken = TokenConversion.ConvertToString(binaryHeader.GetPublicKeyToken());
            string runtimeVersion = addinAssembly.ImageRuntimeVersion;
            string[] classesRoot = null;
            switch (mode)
            {
                case SingletonSettings.UnRegisterMode.Auto:
                    classesRoot = new string[] { "HKEY_CLASSES_ROOT", "HKEY_CURRENT_USER\\Software\\Classes" };
                    break;
                case SingletonSettings.UnRegisterMode.System:
                    classesRoot = new string[] { "HKEY_CLASSES_ROOT" };
                    break;
                case SingletonSettings.UnRegisterMode.User:
                    classesRoot = new string[] { "HKEY_CURRENT_USER\\Software\\Classes" };
                    break;
                default:
                    throw new IndexOutOfRangeException("mode");
            }
            string progid = AttributeReflection.GetAttribute<ProgIdAttribute>(addinClassAttributes).Value;
            string fullClassName = addinClassType.FullName;
            string id = AttributeReflection.GetAttribute<GuidAttribute>(addinClassAttributes).Value;
            string codebase = addinAssembly.CodeBase;

            AddinClassInformations result = new AddinClassInformations(
                assemblyName, assemblyVersion, assemblyCulture, assemblyToken,
                runtimeVersion, classesRoot, progid, fullClassName, id, codebase);

            return result;
        }

        internal static AddinClassInformations Create(Assembly addinAssembly, IEnumerable<object> assemblyAttributes,
                        SingletonSettings.RegisterMode mode, Type addinClassType, IEnumerable<object> addinClassAttributes)
        {
            AssemblyName binaryHeader = addinAssembly.GetName();
            string assemblyName = binaryHeader.Name;
            string assemblyVersion = binaryHeader.Version.ToString();
            string assemblyCulture = CultureInfoConversion.ConvertToString(binaryHeader.CultureInfo);
            string assemblyToken = TokenConversion.ConvertToString(binaryHeader.GetPublicKeyToken());
            string runtimeVersion = addinAssembly.ImageRuntimeVersion;
            string[] classesRoot = new string[] { mode == SingletonSettings.RegisterMode.System ? "HKEY_CLASSES_ROOT" : "HKEY_CURRENT_USER\\Software\\Classes" };
            string progid = AttributeReflection.GetAttribute<ProgIdAttribute>(addinClassAttributes).Value;
            string fullClassName = addinClassType.FullName;
            string id = AttributeReflection.GetAttribute<GuidAttribute>(addinClassAttributes).Value;
            string codebase = addinAssembly.CodeBase;

            AddinClassInformations result = new AddinClassInformations(
                assemblyName, assemblyVersion, assemblyCulture, assemblyToken,
                runtimeVersion, classesRoot, progid, fullClassName, id, codebase);

            return result;
        }

        internal string AssemblyName { get; private set; }

        internal string AssemblyVersion { get; private set; }

        internal string AssemblyCulture{ get; private set; }

        internal string AssemblyToken { get; private set; }

        internal string RuntimeVersion { get; private set; }

        internal string[] ClassesRoot { get; private set; }

        internal string ProgId { get; private set; }

        internal string FullClassName { get; private set; }

        internal string Id { get; private set; }

        internal string ComponentCategoryId { get; private set; }

        internal string Codebase { get; private set; }
    }
}
