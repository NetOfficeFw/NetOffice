using System;
using System.ComponentModel;

namespace NetOffice.Exceptions
{
    /// <summary>
    /// Specify an invoke operation kind
    /// </summary>
    internal enum CallType
    {
        /// <summary>
        /// PropertyGet
        /// </summary>
        PropertyGet = 0,

        /// <summary>
        /// PropertySet
        /// </summary>
        PropertySet = 1,

        /// <summary>
        /// Method or MethodReturn
        /// </summary>
        Method = 2
    }

    internal static class ExceptionMessageBuilder
    {
        /// <summary>
        /// Get exception message based on associated settings
        /// </summary>
        /// <param name="throwedException">exception as any</param>
        /// <param name="instance">caller instance</param>
        /// <param name="name"></param>
        /// <param name="type">name of invoke target</param>
        /// <param name="arguments">arguments as any</param>
        /// <returns>exception message</returns>
        internal static string GetExceptionMessage(Exception throwedException, object instance, string name, CallType type, object[] arguments = null)
        {
            if (null == throwedException)
                throw new ArgumentNullException("throwedException", "<Please report this error.>");

            ICOMObject comObject = instance as ICOMObject;
            if (null == comObject)
                return Settings.Default.ExceptionDefaultMessage;
            Settings settings = comObject.Settings;

            switch (comObject.Settings.ExceptionMessageBehavior)
            {
                case ExceptionMessageHandling.Diagnostics:
                    return GetExceptionDiagnosticsMessage(comObject, name, type, arguments);
                case ExceptionMessageHandling.Default:
                    return settings.ExceptionDefaultMessage;
                case ExceptionMessageHandling.CopyInnerExceptionMessageToTopLevelException:
                    return GetExceptionInnerExceptionMessageToTopLevelMessage(throwedException);
                case ExceptionMessageHandling.CopyAllInnerExceptionMessagesToTopLevelException:
                    return GetExceptionAllInnerExceptionMessagesToTopLevelMessage(throwedException);
                case ExceptionMessageHandling.DiagnosticsAndInnerMessage:
                    return GetExceptionDiagnosticsInnerMessage(throwedException, comObject, name, type, arguments);
                default:
                    throw new NetOfficeException("<Unexpected exception behavior. Please report this error.>");
            }
        }

        private static string GetExceptionDiagnosticsInnerMessage(Exception throwedException, ICOMObject comObject, string name, CallType type, object[] arguments = null)
        {
            string diagMessage = GetExceptionDiagnosticsMessage(comObject, name, type, arguments);
            string message = throwedException.Message;
            while (throwedException.InnerException != null)
            {
                message = throwedException.InnerException.Message;
                throwedException = throwedException.InnerException;
            }
            return diagMessage + "{" + message + "}";
        }

        /// <summary>
        /// Get diagnostic exception message
        /// </summary>
        /// <param name="comObject">caller instance</param>
        /// <param name="name">name of invoke target</param>
        /// <param name="type">type of invoke target</param>
        /// <param name="arguments">arguments as any</param>
        /// <returns>diagnostic exception message or error message if an exception occurs</returns>
        private static string GetExceptionDiagnosticsMessage(ICOMObject comObject, string name, CallType type, object[] arguments = null)
        {
            try
            {
                Settings settings = comObject.Settings;
                string diagMessage = settings.ExceptionDiagnosticsMessage;
                if (String.IsNullOrWhiteSpace(diagMessage))
                    return diagMessage;

                diagMessage = diagMessage.Replace("{CallType}", type.ToString());
                diagMessage = diagMessage.Replace("{CallInstance}", comObject.InstanceFriendlyName);
                diagMessage = diagMessage.Replace("{Name}", name);
                if (diagMessage.IndexOf("{Args}") > -1)
                {
                    string argsString = String.Empty;
                    if (null != arguments && arguments.Length > 0)
                    {
                        for (int i = 0; i < arguments.Length; i++)
                        {
                            object arg = argsString[i];
                            if (null != argsString)
                            {
                                if (arg == Type.Missing)
                                    argsString += "<Type.Missing>";
                                else
                                {
                                    ICOMObject comObjectArg = arg as ICOMObject;
                                    if (null != comObjectArg)
                                    {
                                        argsString += comObjectArg.InstanceFriendlyName;
                                    }
                                    else if (arg is MarshalByRefObject)
                                    {
                                        argsString += TryGetProxyClassName(arg);
                                    }
                                    else
                                    {
                                        argsString += arg.ToString();
                                    }
                                }
                            }
                            else
                                argsString += "<null>";

                            if (i < arguments.Length - 1)
                                argsString += ", ";
                        }
                    }
                    diagMessage = diagMessage.Replace("{Args}", argsString);
                }

                return diagMessage;
            }
            catch
            {
                return "<Failed to create exception message. Please report this error.>";
            }
        }

        /// <summary>
        /// Get associated settings default message
        /// </summary>
        /// <param name="comObject">caller instance</param>
        /// <returns>default exception message</returns>
        private static string GetExceptionDefaultMessage(ICOMObject comObject)
        {
            return comObject.Settings.ExceptionDefaultMessage;
        }

        /// <summary>
        /// Get most inner/bottom exception message
        /// </summary>
        /// <param name="throwedException">exception as any</param>
        /// <returns>most inner exception message</returns>
        private static string GetExceptionInnerExceptionMessageToTopLevelMessage(Exception throwedException)
        {
            string message = throwedException.Message;
            while (throwedException.InnerException != null)
            {
                message = throwedException.InnerException.Message;
                throwedException = throwedException.InnerException;
            }
            return message;
        }

        /// <summary>
        /// Get all exception/inner exception messages as summary
        /// </summary>
        /// <param name="throwedException">exception</param>
        /// <returns>exception message summary</returns>
        private static string GetExceptionAllInnerExceptionMessagesToTopLevelMessage(Exception throwedException)
        {
            string messageSummary = throwedException.Message + Environment.NewLine;
            while (throwedException.InnerException != null)
            {
                messageSummary += throwedException.InnerException.Message + Environment.NewLine;
                throwedException = throwedException.InnerException;
            }
            return messageSummary;
        }

        /// <summary>
        /// Get default settings exception message
        /// </summary>
        /// <returns></returns>
        internal static string GetDefaultExceptionMessage()
        {
            return Settings.Default.ExceptionDefaultMessage;
        }

        /// <summary>
        /// Try get proxy class and suspress any exception
        /// </summary>
        /// <param name="proxy">proxy instance</param>
        /// <returns>class name or System._ComObject</returns>
        private static string TryGetProxyClassName(object proxy)
        {
            try
            {
                return TypeDescriptor.GetClassName(proxy);
            }
            catch
            {
                return "System._ComObject";
            }
        }
    }
}