using System;
using System.Runtime.InteropServices;
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
        /// <summary>
        /// Analyze a COMAddin for the TweakAttribute and try to set given arguments(registry) if exists
        /// </summary>
        /// <param name="factory">current used factory or null for default</param>
        /// <param name="addinType">Type info from COMAddin instance</param>
        /// <param name="registryEndPoint">specific office registry key endpoint</param>
        public static void EnableTweaks(Core factory, Type addinType, string registryEndPoint)
        {
            try
            {
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
                    TweakConsoleMode(factory, key);
                    TweakSharedOutput(factory, key);
                    TweakAddHocLoading(factory, key);
                    TweakDeepLoading(factory, key);
                    TweakDebugOutput(factory, key);
                    TweakExceptionHandling(factory, key);
                    TweakExceptionMessage(factory, key);
                    TweakThreadCulture(factory, key);
                    TweakMessageFilter(factory, key);
                    TweakSafeMode(factory, key);
                    TweakEventOutput(factory, key);

                    key.Close();
                    key.Dispose();
                }
                hiveKey.Close();
                hiveKey.Dispose();
            }
            catch (Exception exception)
            {
                factory.Console.WriteException(exception);
            }
        }

        private static void TweakConsoleMode(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOConsoleMode", null) as string;
            if (null != value)
            {
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

        private static void TweakSharedOutput(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOConsoleShare", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Console.EnableSharedOutput = true;
                else
                    factory.Console.EnableSharedOutput = false;
            }
        }

        private static void TweakExceptionHandling(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOExceptionHandling", null) as string;
            if (null != value)
            {
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


        private static void TweakExceptionMessage(Core factory, RegistryKey key)
        {
             string value = key.GetValue("NOExceptionMessage", null) as string;
             if (null != value)
             {
                 factory.Settings.ExceptionMessage = value;
             }
        }

        private static void TweakThreadCulture(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOCultureInfo", null) as string;
            if (null != value)
            {
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

        private static void TweakMessageFilter(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOMessageFilter", null) as string;
            if (null != value)
            {
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

        private static void TweakSafeMode(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOSafeMode", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableSafeMode = true;
                else
                    factory.Settings.EnableSafeMode = false;
            }
        }

        private static void TweakAddHocLoading(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOAdHocLoad", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableAdHocLoading = true;
                else
                    factory.Settings.EnableAdHocLoading = false;
            }
        }

        private static void TweakDeepLoading(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NODeepLoad", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableDeepLoading = true;
                else
                    factory.Settings.EnableDeepLoading = false;
            }
        }

        private static void TweakDebugOutput(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NODebugOut", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableDebugOutput = true;
                else
                    factory.Settings.EnableDebugOutput = false;
            }
        }

        private static void TweakEventOutput(Core factory, RegistryKey key)
        {
            string value = key.GetValue("NOEventOut", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    factory.Settings.EnableEventDebugOutput = true;
                else
                    factory.Settings.EnableEventDebugOutput = false;
            }
        }
    }
}
