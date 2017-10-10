using System.Collections.Generic;
using NetRuntimeSystem = System;
using Microsoft.Win32;

namespace NetOffice.WordApi.Tools
{
    /// <summary>
    /// Specify a Document Inspector
    /// PLEASE NOTE: Document Inspector must register in MS Word LocalMachine Registry Hive Key
    /// </summary>
    [NetRuntimeSystem.AttributeUsage(NetRuntimeSystem.AttributeTargets.Class, AllowMultiple = false)]
    public class DocumentInspectorAttribute : NetRuntimeSystem.Attribute
    {
        private static string _targetKeyTemplate = @"Software\Microsoft\Office\{0}\{1}\Document Inspectors\{2}";

        /// <summary>
        /// Name of the Document Inspector
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// Description of the Document Inspector
        /// </summary>
        public readonly string Description;

        /// <summary>
        /// Target Application Versions by comma delimiter
        /// 12,14,15,16 for all versions
        /// </summary>
        private string ApplicationVersion;

        /// <summary>
        /// 1 = Document Inspector is selected
        /// </summary>
        public readonly int Selected;

        /// <summary>
        /// Pre-processed ApplicationVersion;
        /// </summary>
        public readonly double[] ProcessedApplicationVersion;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">name of the document inspector</param>
        /// <param name="description">description of the document inspector</param>
        /// <param name="applicationVersion">target application versions by comma delimiter</param>
        /// <param name="selected">document inspector is selected</param>
        public DocumentInspectorAttribute(string name, string description, string applicationVersion, int selected)
        {
            Name = name;
            Description = description;
            ApplicationVersion = applicationVersion;
            Selected = selected;
            ApplicationVersion = applicationVersion;

            List<double> list = new List<double>();
            string[] versions = ApplicationVersion.Split(new string[] { "," }, NetRuntimeSystem.StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in versions)
            {
                double i = 0;
                if (double.TryParse(item, out i))
                    list.Add(i);
            }
            ProcessedApplicationVersion = list.ToArray();
        }

        /// <summary>
        /// Reflect a type and returns DocumentInspectorAttribute instances
        /// </summary>
        /// <param name="type">type to reflect</param>
        /// <returns>DocumentInspectorAttribute instances</returns>
        public static DocumentInspectorAttribute[] GetAttributes(NetRuntimeSystem.Type type)
        {
            object[] attributes = type.GetCustomAttributes(typeof(DocumentInspectorAttribute), true);
            DocumentInspectorAttribute[] result = new DocumentInspectorAttribute[attributes.Length];
            for (int i = 0; i < attributes.Length; i++)
                result[i] = attributes[i] as DocumentInspectorAttribute;
            return result;
        }

        /// <summary>
        /// Create DocumentInspector Registry Key
        /// </summary>
        /// <param name="officeProduct">office application name</param>
        /// <param name="name">inspector name</param>
        /// <param name="version">word version</param>
        /// <param name="selected">inspector is selected by default</param>
        /// <param name="typeid">addin clsid</param>
        public static void CreateKey(string officeProduct, string name, double version, int selected, string typeid)
        {
            string targetKey = string.Format(_targetKeyTemplate, ValidateVersion(version), officeProduct, name);
            RegistryKey applicationKey = null;
            applicationKey = Registry.LocalMachine.CreateSubKey(targetKey);
            applicationKey.Close();

            applicationKey = Registry.LocalMachine.OpenSubKey(targetKey, true);

            if (!typeid.StartsWith("{"))
                typeid = "{" + typeid;
            if (!typeid.EndsWith("}"))
                typeid = typeid + "}";

            applicationKey.SetValue("CLSID", typeid, RegistryValueKind.String);
            applicationKey.SetValue("Selected", selected, RegistryValueKind.DWord);
            applicationKey.Close();
        }

        /// <summary>
        /// Try delete DocumentInspector Registry Key
        /// </summary>
        /// <param name="officeProduct">office application name</param>
        /// <param name="name">inspector name</param>
        /// <param name="version">word version</param>
        public static void TryDeleteKey(string officeProduct, string name, double version)
        {
            string targetKey = string.Format(_targetKeyTemplate, ValidateVersion(version), officeProduct, name);
            Registry.LocalMachine.DeleteSubKey(targetKey, false);
        }

        private static string ValidateVersion(double version)
        {
            string result = version.ToString();
            if (result.IndexOf(".") == -1 && false == result.EndsWith(".0"))
                result += ".0";
            return result;
        }
    }
}