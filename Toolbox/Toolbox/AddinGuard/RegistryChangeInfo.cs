using System;
using System.Collections.Generic;
using Microsoft.Win32;
using System.Text;

namespace NetOffice.DeveloperToolbox.AddinGuard
{
    class AddinValuesIncrementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
    }

    class AddinValuesDecrementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
    }

    class AddinValueNameChangedInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string NewValueName { get; set; }
        public string OldValueName { get; set; }
    }

    class AddinValueKindChangedInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
        public RegistryValueKind NewValueKind { get; set; }
        public RegistryValueKind OldValueKind { get; set; }
    }

    class AddinValueValueChangedInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
        public object NewValue { get; set; }
        public object OldValue { get; set; }
    }

    class AddinValueValueRestoredInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
        public string ValueName { get; set; }
        public object OldValue { get; set; }
        public object RestoredValue { get; set; }
    }

    class NewDeactivatedElementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string Name { get; set; }
        public string OfficeProductVersion { get; set; }
    }

    class DeleteDeactivatedElementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string Name { get; set; }
        public string OfficeProductVersion { get; set; }
    }

    class AddinSubkeysIncrementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
    }

    class AddinSubkeysDecrementInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string KeyName { get; set; }
    }

    class AddinSubkeyNameChangedInfo : RegistryChangeInfo
    {
        public RegistryKey RootKey { get; set; }
        public string KeyPath { get; set; }
        public string OldKeyName { get; set; }
        public string NewKeyName { get; set; }
    }

    class RegistryChangeInfo
    {
      
    }

}
