using System;
using System.Collections.Generic;
using Microsoft.Win32;
using NetOffice.OutlookApi.Enums;

namespace NetOffice.OutlookApi.Tools
{
    /// <summary>
    /// Specify a provided FormRegion
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class FormRegionAttribute : System.Attribute
    {
        /// <summary>
        /// Name of the FormRegion
        /// </summary>
        public readonly string Name;

        /// <summary>
        /// Category of the FormRegion
        /// </summary>
        public readonly string Category;

        /// <summary>
        /// Manifest Resource Path of the FormRegion
        /// </summary>
        public readonly string ManifestFile;

        /// <summary>
        /// Storage Resource Path of the FormRegion
        /// </summary>
        public readonly string StorageFile;

        /// <summary>
        /// Target Localization ID
        /// </summary>
        public readonly int LCID;

        /// <summary>
        /// Given Icon Range
        /// </summary>
        public readonly OlFormRegionIcon OlIcon;

        /// <summary>
        /// Given Icon is generaly used
        /// </summary>
        public bool OlIconWildcard;

        /// <summary>
        /// Icon Resource Path of the FormRegion
        /// </summary>
        public readonly string IconFile;

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">name of the formregion</param>
        /// <param name="category">category of the formregion</param>
        /// <param name="manifestFile">manifest Resource path of the formregion</param>
        /// <param name="storageFile">storage resource path of the formregion</param>
        /// <param name="olIcon">given icon range</param>
        /// <param name="iconFile">icon resource path of the formregion</param>
        /// <param name="lcid">target localization id</param>
        public FormRegionAttribute(string name, string category, string manifestFile, string storageFile, OlFormRegionIcon olIcon, string iconFile, int lcid)
        {
            Name = name;
            Category = category;
            ManifestFile = manifestFile;
            StorageFile = storageFile;
            OlIcon = olIcon;
            IconFile = iconFile;
            LCID = lcid;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">name of the formregion</param>
        /// <param name="category">category of the formregion</param>
        /// <param name="manifestFile">manifest Resource path of the formregion</param>
        /// <param name="storageFile">storage resource path of the formregion</param>
        /// <param name="olIcon">given icon range</param>
        /// <param name="iconFile">icon resource path of the formregion</param>
        public FormRegionAttribute(string name, string category, string manifestFile, string storageFile, OlFormRegionIcon olIcon, string iconFile)
        {
            Name = name;
            Category = category;
            ManifestFile = manifestFile;
            StorageFile = storageFile;
            OlIcon = olIcon;
            IconFile = iconFile;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">name of the formregion</param>
        /// <param name="category">category of the formregion</param>
        /// <param name="manifestFile">manifest Resource path of the formregion</param>
        /// <param name="storageFile">storage resource path of the formregion</param>
        /// <param name="iconFile">icon resource path of the formregion</param>
        /// <param name="lcid">target localization id</param>
        public FormRegionAttribute(string name, string category, string manifestFile, string storageFile, string iconFile, int lcid)
        {
            Name = name;
            Category = category;
            ManifestFile = manifestFile;
            StorageFile = storageFile;
            OlIconWildcard = true;
            IconFile = iconFile;
            LCID = lcid;
        }

        /// <summary>
        /// Creates an instance of the attribute
        /// </summary>
        /// <param name="name">name of the formregion</param>
        /// <param name="category">category of the formregion</param>
        /// <param name="manifestFile">manifest Resource path of the formregion</param>
        /// <param name="storageFile">storage resource path of the formregion</param>
        /// <param name="iconFile">icon resource path of the formregion</param>
        public FormRegionAttribute(string name, string category, string manifestFile, string storageFile, string iconFile)
        {
            Name = name;
            Category = category;
            ManifestFile = manifestFile;
            StorageFile = storageFile;
            OlIconWildcard = true;
            IconFile = iconFile;
        }

        /// <summary>
        /// Get attribute by given arguments
        /// </summary>
        /// <param name="type">type to analyze</param>
        /// <param name="formRegionName">target formregion name</param>
        /// <param name="lcid">localization id</param>
        /// <returns>attribute or null</returns>
        public static FormRegionAttribute GetAttribute(Type type, string formRegionName, int lcid)
        {
            object[] attributes = type.GetCustomAttributes(typeof(FormRegionAttribute), true);
            foreach (FormRegionAttribute item in attributes)
            {
                if (item.LCID != 0 && item.LCID != lcid)
                    continue;
                if (item.Name == formRegionName)
                    return item;
            }
            return null;
        }

        /// <summary>
        /// Get FormRegion attributes
        /// </summary>
        /// <param name="type">type to analyze</param>
        /// <returns>attributes</returns>
        public static IEnumerable<FormRegionAttribute> GetAttributes(Type type)
        {
            object[] attributes = type.GetCustomAttributes(typeof(FormRegionAttribute), true);
            FormRegionAttribute[] result = new FormRegionAttribute[attributes.Length];
            for (int i = 0; i < attributes.Length; i++)
                result[i] = attributes[i] as FormRegionAttribute;
            return result;
        }

        /// <summary>
        /// Create FormRegion registry value
        /// </summary>
        /// <param name="isSystem">system or current user</param>
        /// <param name="officeKey">application key - only outlook is supported</param>
        /// <param name="progId">addin progid</param>
        /// <param name="category">region category</param>
        /// <param name="name">region name</param>
        public static void CreateKey(bool isSystem, string officeKey, string progId, string category, string name)
        {
            string targetKey = officeKey + category;
            RegistryKey applicationKey = null;
            if (isSystem)
                applicationKey = Registry.LocalMachine.CreateSubKey(targetKey);
            else
                applicationKey = Registry.CurrentUser.CreateSubKey(targetKey);

            applicationKey.Close();

            if (isSystem)
                applicationKey = Registry.LocalMachine.OpenSubKey(targetKey, true);
            else
                applicationKey = Registry.CurrentUser.OpenSubKey(targetKey, true);

            applicationKey.SetValue(name , "=" + progId, RegistryValueKind.String);
            applicationKey.Close();
        }

        /// <summary>
        /// Delete FormRegion registry value
        /// </summary>
        /// <param name="isSystem">system or current user</param>
        /// <param name="officeKey">application key - only outlook is supported</param>
        /// <param name="progId">addin progid</param>
        /// <param name="category">region category</param>
        /// <param name="name">region name</param>
        public static bool TryDeleteKey(bool isSystem, string officeKey, string progId, string category, string name)
        {
            try
            {
                string targetKey = officeKey + category;
                RegistryKey key = null;
                if (isSystem)
                    key = Registry.LocalMachine.OpenSubKey(targetKey, true);
                else
                    key = Registry.CurrentUser.OpenSubKey(targetKey, true);

                key.DeleteValue(name);
                key.Close();

                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}