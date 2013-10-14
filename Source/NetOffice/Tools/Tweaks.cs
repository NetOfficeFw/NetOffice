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
        /// <param name="addinType">Type info from COMAddin instance</param>
        /// <param name="registryEndPoint">specific office registry key endpoint</param>
        public static void EnableTweaks(Type addinType, string registryEndPoint)
        {
            try
            {
                if (null == addinType)
                    return;

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

                RegistryKey key = hiveKey.OpenSubKey("Software\\Microsoft\\Office\\" + registryEndPoint + "\\" + progIDAttribute.Value);
                if (null != key)
                {
                    TweakConsoleMode(key);
                    TweakSharedOutput(key);
                    TweakExceptionHandling(key);
                    TweakExceptionMessage(key);
                    TweakThreadCulture(key);
                    TweakMessageFilter(key);
                    TweakSafeMode(key);
                    TweakThreadSafe(key);
                    TweakAddHocLoading(key);
                    TweakDeepLoading(key);
                    TweakDebugOutput(key);
                    TweakEventOutput(key);

                    key.Close();
                    key.Dispose();
                }
                hiveKey.Close();
                hiveKey.Dispose();
            }
            catch (Exception exception)
            {
                DebugConsole.WriteException(exception);
            }
        }

        private static void TweakConsoleMode(RegistryKey key)
        {
            string value = key.GetValue("NOConsoleMode", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "none":
                        DebugConsole.Mode = ConsoleMode.None;
                        return;
                    case "console":
                        DebugConsole.Mode = ConsoleMode.Console;
                        return;
                    case "trace":
                        DebugConsole.Mode = ConsoleMode.Trace;
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
                        DebugConsole.FileName = logFile;
                        DebugConsole.Mode = ConsoleMode.LogFile;
                    }
                }
            }
        }

        private static void TweakSharedOutput(RegistryKey key)
        {
            string value = key.GetValue("NOConsoleShare", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    DebugConsole.EnableSharedOutput = true;
                else
                    DebugConsole.EnableSharedOutput = false;
            }
        }

        private static void TweakExceptionHandling(RegistryKey key)
        {
            string value = key.GetValue("NOExceptionHandling", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "default":
                        Settings.UseExceptionMessage = ExceptionMessageHandling.Default;
                        return;
                    case "copyInnerexceptionmessagetotoplevelexception":
                        Settings.UseExceptionMessage = ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException;
                        return;
                    case "copyallinnerexceptionmessagestotoplevelexception":
                        Settings.UseExceptionMessage = ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException;
                        return;
                    default:
                        break;
                }
            }
        }


        private static void TweakExceptionMessage(RegistryKey key)
        {
             string value = key.GetValue("NOExceptionMessage", null) as string;
             if (null != value)
             {
                 Settings.ExceptionMessage = value;
             }
        }

        private static void TweakThreadCulture(RegistryKey key)
        {
            string value = key.GetValue("NOCultureInfo", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                try
                {
                    Settings.ThreadCulture = System.Globalization.CultureInfo.GetCultureInfo(value);
                }
                catch (Exception exception)
                {
                    DebugConsole.WriteException(exception);
                }               
            }
        }

        private static void TweakMessageFilter(RegistryKey key)
        {
            string value = key.GetValue("NOMessageFilter", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                switch (value)
                {
                    case "immediately":
                        Settings.MessageFilter.RetryMode = RetryMessageFilterMode.Immediately;
                        return;
                    case "delayed":
                        Settings.MessageFilter.RetryMode = RetryMessageFilterMode.Delayed;
                        return;
                    case "None":
                        Settings.MessageFilter.RetryMode = RetryMessageFilterMode.None;
                        return;
                    default:
                        break;
                }
            }
        }

        private static void TweakSafeMode(RegistryKey key)
        {
            string value = key.GetValue("NOSafeMode", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableSafeMode = true;
                else
                    Settings.EnableSafeMode = false;
            }
        }

        private static void TweakThreadSafe(RegistryKey key)
        {
            string value = key.GetValue("NOThreadSafe", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableThreadSafe = true;
                else
                    Settings.EnableThreadSafe = false;
            }
        }

        private static void TweakAddHocLoading(RegistryKey key)
        {
            string value = key.GetValue("NOAdHocLoad", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableAdHocLoading = true;
                else
                    Settings.EnableAdHocLoading = false;
            }
        }

        private static void TweakDeepLoading(RegistryKey key)
        {
            string value = key.GetValue("NODeepLoad", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableDeepLoading = true;
                else
                    Settings.EnableDeepLoading = false;
            }
        }

        private static void TweakDebugOutput(RegistryKey key)
        {
            string value = key.GetValue("NODebugOut", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableDebugOutput = true;
                else
                    Settings.EnableDebugOutput = false;
            }
        }

        private static void TweakEventOutput(RegistryKey key)
        {
            string value = key.GetValue("NOEventOut", null) as string;
            if (null != value)
            {
                value = value.ToLower().Trim();
                if (value == "enabled")
                    Settings.EnableEventDebugOutput = true;
                else
                    Settings.EnableEventDebugOutput = false;
            }
        }
    }
}
