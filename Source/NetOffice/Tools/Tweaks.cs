#define RegKeyDisposeAvailable

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Reflection;
using Microsoft.Win32;
using System.Text;
using NetOffice.Tools;

namespace NetOffice.Tools
{ 
    /// <summary>
    /// Tweak Handler to customize some settings at runtime (if wanted)
    /// </summary>
    public static class Tweaks
    {
       

        #region Fields

        private static string[] _noTweakNames = new string[] { "NOConsoleMode", "NOConsoleShare", "NOExceptionHandling", "NOExceptionMessage", "NOCultureInfo", 
                                                               "NOMessageFilter", "NOSafeMode", "NOAdHocLoad", "NODeepLoad",
                                                               "NODebugOut", "NODebugOut", "NOEventOut" };
        private static string[] _addinValueNames = new string[] { "FriendlyName", "Description", "LoadBehavior", "CommandLineSafe" };

        #endregion

        #region Ctor

        /// <summary>
        /// Creates no instance of the class
        /// </summary>
        static Tweaks()
        {
            CustomTweaks = new Dictionary<int, Dictionary<string, string>>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Store applied custom teaks. int = GetHashCode() from COMAddin instance. Dictionary string string = name, value of applied tweak
        /// </summary>
        private static Dictionary<int, Dictionary<string, string>> CustomTweaks { get; set; }

        #endregion

        /// <summary>
        /// Analyze a COMAddin for the TweakAttribute and try to set given arguments(registry) if exists
        /// </summary>
        /// <param name="factory">current used factory or null for default</param>
        /// <param name="addinInstance">COMAddin instance</param>
        /// <param name="addinType">Type info from COMAddin instance</param>
        /// <param name="registryEndPoint">specific office registry key endpoint</param>
        public static void ApplyTweaks(Core factory, object addinInstance, Type addinType, string registryEndPoint)
        {
            try
            {
                if (null == addinInstance)
                    return;

                if (null == factory)
                    factory = Core.Default;

                TweakAttribute tweakAttribute = AttributeHelper.GetTweakAttribute(addinType);
                if (null == tweakAttribute || false == tweakAttribute.Enabled)
                    return;

                ProgIdAttribute progIDAttribute = AttributeHelper.GetProgIDAttribute(addinType);
                if (null == progIDAttribute)
                    return;

                RegistryKey hiveKey = Registry.CurrentUser;
                RegistryLocationAttribute locationAttribute = AttributeHelper.GetRegistryLocationAttribute(addinType);
                if (null != locationAttribute && locationAttribute.Value == RegistrySaveLocation.LocalMachine)
                    hiveKey = Registry.LocalMachine;

                RegistryKey key = hiveKey.OpenSubKey("Software\\Microsoft\\Office\\" + registryEndPoint + "\\Addins\\" + progIDAttribute.Value);
                if (null != key)
                {
                    TweakProxyCountChannel(factory, addinInstance, addinType, key);
                    TweakConsoleMode(factory, addinInstance, addinType, key);
                    TweakSharedOutput(factory, addinInstance, addinType, key);
                    TweakAddHocLoading(factory, addinInstance, addinType, key);
                    TweakDeepLoading(factory, addinInstance, addinType, key);
                    TweakDebugOutput(factory, addinInstance, addinType, key);
                    TweakExceptionHandling(factory, addinInstance, addinType, key);
                    TweakExceptionMessage(factory, addinInstance, addinType, key);
                    TweakThreadCulture(factory, addinInstance, addinType, key);
                    TweakMessageFilter(factory, addinInstance, addinType, key);
                    TweakSafeMode(factory, addinInstance, addinType, key);
                    TweakEventOutput(factory, addinInstance, addinType, key);
                    Dictionary<string, string> customTweaks = ApplyCustomTweaks(factory, addinInstance, addinType, key);
                    AddCustomAppliedTweaks(addinInstance.GetHashCode(), customTweaks);
                    key.Close();
                    // key.Dispose(); not available in previous .net versions (but in fact RegistryKey implements always IDisposable huuu?)
                }
                hiveKey.Close();
                // hiveKey.Dispose();  not available in previous .net versions
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
            }
        }

        /// <summary>
        /// Dispose applied tweaks for COMAddin instance
        /// </summary>
        /// <param name="factory">current used factory or null for default</param>
        /// <param name="addinInstance">COMAddin instance</param>
        /// <param name="addinType">Type info from COMAddin instance</param>
        public static void DisposeTweaks(Core factory, object addinInstance, Type addinType)
        {
            try
            {
                DisposeCustomAppliedTweaks(factory, addinInstance, addinType);
                RemoveCustomAppliedTweaks(factory, addinInstance, addinType);
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
            }
        }

        #region Custom Tweaks

        /// <summary>
        /// Returns info the regkey value name is addin standard or NetOffice tweak
        /// </summary>
        /// <param name="name">target regky name</param>
        /// <returns>true if found</returns>
        private static bool IsWellKnownName(string name)
        {
            foreach (string item in _noTweakNames)
            {
                if (name.Equals(item, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }

            foreach (string item in _addinValueNames)
            {
                if (name.Equals(item, StringComparison.InvariantCultureIgnoreCase))
                    return true;
            }
            return false;
        }

        private static Dictionary<string, string> ApplyCustomTweaks(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            string[] names = key.GetValueNames();
            foreach (string item in names)
            {
                if (IsWellKnownName(item))
                    continue;
                string value = key.GetValue(item, null) as string;
                if (null != value)
                {
                    if (CallAllowApplyTweak(factory, addinInstance, addinType, item, value))
                    {
                        CallApplyCustomTweak(factory, addinInstance, addinType, item, value);
                        result.Add(item, value);
                    }
                }
            }
            return result;
        }

        private static void DisposeCustomAppliedTweaks(Core factory, object addinInstance, Type addinType)
        {
            if (CustomTweaks.ContainsKey(addinInstance.GetHashCode()))
            {
                Dictionary<string, string> customTeaks = CustomTweaks[addinInstance.GetHashCode()];
                foreach (var item in customTeaks)
                    CallDisposeCustomTweak(factory, addinInstance, addinType, item.Key, item.Value);
            }
        }

        private static void AddCustomAppliedTweaks(int hashCode, Dictionary<string, string> customTweaks)
        {
            if (CustomTweaks.ContainsKey(hashCode))
                CustomTweaks[hashCode] = customTweaks;
            else
                CustomTweaks.Add(hashCode, customTweaks);
        }

        private static void RemoveCustomAppliedTweaks(Core factory, object addinInstance, Type addinType)
        {
            if (CustomTweaks.ContainsKey(addinInstance.GetHashCode()))
                CustomTweaks.Remove(addinInstance.GetHashCode());
        }

        #endregion

        #region Caller Methods

        private static bool CallAllowApplyTweak(Core factory, object addinInstance, Type addinType, string name, string value)
        {
            try
            {
                if (null == addinInstance)
                    return false;
                if (null == addinType)
                    return false;
                return (bool)addinType.InvokeMember("AllowApplyTweak", BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.InvokeMethod, null, addinInstance, new object[] { name, value });
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
                return false;
            }
        }

        private static void CallApplyCustomTweak(Core factory, object addinInstance, Type addinType, string name, string value)
        {
            try
            {
                if (null == addinInstance)
                    return;
                if (null == addinType)
                    return;
                addinType.InvokeMember("ApplyCustomTweak", BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.InvokeMethod, null, addinInstance, new object[] { name, value });
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
            }
        }

        private static void CallDisposeCustomTweak(Core factory, object addinInstance, Type addinType, string name, string value)
        {
            try
            {
                if (null == addinInstance)
                    return;
                if (null == addinType)
                    return;
                addinType.InvokeMember("DisposeCustomTweak", BindingFlags.DeclaredOnly | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.InvokeMethod, null, addinInstance, new object[] { name, value });
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
            }
        }

        #endregion

        #region NO Tweaks

        private static void TweakProxyCountChannel(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOEnableProxyCountChannel", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOEnableProxyCountChannel", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value.StartsWith("enabled", StringComparison.InvariantCultureIgnoreCase))
                {
                    int pos = value.IndexOf(";", StringComparison.InvariantCultureIgnoreCase);
                    if (pos > -1)
                    {
                        string channelName = value.Substring(pos + 1);
                        factory.Settings.EnableProxyCountChannel = true;
                        factory.Settings.ProxyCountChannelName = channelName;
                    }
                }
            }
        }

        private static void TweakConsoleMode(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {

            string value = key.GetValue("NOConsoleMode", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOConsoleMode", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "none":
                        factory.Console.Mode = DebugConsoleMode.None;
                        return;
                    case "console":
                        factory.Console.Mode = DebugConsoleMode.Console;
                        return;
                    case "trace":
                        factory.Console.Mode = DebugConsoleMode.Trace;
                        return;
                    default:
                        break;
                }

                if (value.StartsWith("logfile", StringComparison.InvariantCultureIgnoreCase))
                {
                    int pos = value.IndexOf(";", StringComparison.InvariantCultureIgnoreCase);
                    if (pos > -1)
                    {
                        string logFile = value.Substring(pos + 1);
                        factory.Console.FileName = logFile;
                        factory.Console.Mode = DebugConsoleMode.LogFile;
                    }
                }
            }
        }

        private static void TweakSharedOutput(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOConsoleShare", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOConsoleShare", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Console.EnableSharedOutput = true;
                else
                    factory.Console.EnableSharedOutput = false;
            }
        }

        private static void TweakExceptionHandling(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOExceptionHandling", null) as string;
            if (null != value)
            {

                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOExceptionHandling", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "default":
                        factory.Settings.UseExceptionMessage = ExceptionMessageHandling.Default;
                        return;
                    case "copyInnerexceptionmessagetotoplevelexception":
                        factory.Settings.UseExceptionMessage = ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException;
                        return;
                    case "copyallinnerexceptionmessagestotoplevelexception":
                        factory.Settings.UseExceptionMessage = ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException;
                        return;
                    default:
                        break;
                }
            }
        }

        private static void TweakExceptionMessage(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOExceptionMessage", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOExceptionMessage", value);
                if (!allow)
                    return;
                factory.Settings.ExceptionMessage = value;
            }
        }

        private static void TweakThreadCulture(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOCultureInfo", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOCultureInfo", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                try
                {
                    factory.Settings.ThreadCulture = System.Globalization.CultureInfo.GetCultureInfo(value);
                }
                catch (Exception exception)
                {
                    factory.Console.WriteException(exception);
                }
            }
        }

        private static void TweakMessageFilter(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOMessageFilter", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOMessageFilter", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "immediately":
                        factory.Settings.MessageFilter.RetryMode = RetryMessageFilterMode.Immediately;
                        return;
                    case "delayed":
                        factory.Settings.MessageFilter.RetryMode = RetryMessageFilterMode.Delayed;
                        return;
                    case "None":
                        factory.Settings.MessageFilter.RetryMode = RetryMessageFilterMode.None;
                        return;
                    default:
                        break;
                }
            }
        }

        private static void TweakSafeMode(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOSafeMode", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOSafeMode", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableSafeMode = true;
                else
                    factory.Settings.EnableSafeMode = false;
            }
        }

        private static void TweakAddHocLoading(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOAdHocLoad", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOAdHocLoad", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableAdHocLoading = true;
                else
                    factory.Settings.EnableAdHocLoading = false;
            }
        }

        private static void TweakDeepLoading(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NODeepLoad", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NODeepLoad", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableDeepLoading = true;
                else
                    factory.Settings.EnableDeepLoading = false;
            }
        }

        private static void TweakDebugOutput(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NODebugOut", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NODebugOut", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableDebugOutput = true;
                else
                    factory.Settings.EnableDebugOutput = false;
            }
        }

        private static void TweakEventOutput(Core factory, object addinInstance, Type addinType, RegistryKey key)
        {
            string value = key.GetValue("NOEventOut", null) as string;
            if (null != value)
            {
                bool allow = CallAllowApplyTweak(factory, addinInstance, addinType, "NOEventOut", value);
                if (!allow)
                    return;
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableEventDebugOutput = true;
                else
                    factory.Settings.EnableEventDebugOutput = false;
            }
        }

        #endregion
    }
}
