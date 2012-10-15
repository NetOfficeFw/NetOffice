																																						  
using System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.InteropServices;

using NetOffice;
using NetOffice.Tools;
using Visio = NetOffice.VisioApi;

namespace NetOffice.VisioApi.Tools
{
    /// <summary>
    /// The class provides a lot of essential functionality for an MS-Project COMAddin
    /// </summary>
    public abstract class COMAddin : IDTExtensibility2
    {
        #region Fields

        /// <summary>
        /// MS-Visio Registry Path 
        /// </summary>
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\Visio\\AddIns\\";

        #endregion
        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public COMAddin()
        {
            try
            {
                Type = this.GetType();
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if(!handled)
                    throw exception;
            }
        }

        #endregion

        #region Properties

        /// <summary>
        /// Type Information of the instance
        /// </summary>
        protected Type Type { get; set; }

        /// <summary>
        /// Host Application Instance
        /// </summary>
        protected Visio.Application Application { get; private set; }
      
        #endregion

        #region (IDTExtensibility2) Events 

        /// <summary>
        /// The OnStartupComplete event occurs when the host application completes its startup routines, in the case where the COM add-in loads at startup. 
        /// If the add-in is not loaded when the application loads, the OnStartupComplete event does not occur — 
        /// even when the user loads the add-in in the COM Add-ins dialog box. When this event does occur, it occurs after the OnConnection event.
        /// You can use the OnStartupComplete  event procedure to run code that interacts with the application and that should not be run until the application has finished loading. 
        /// For example, if you want to display a form that gives users a choice of documents to create when they start the application, 
        /// you can put that code in the OnStartupComplete event procedure.
        /// </summary>
        public event OnStartupCompleteEventHandler OnStartupComplete;

        /// <summary>
        /// The Shutdown event occurs when the COM add-in is unloaded. 
        /// You can use the OnDisconnection event procedure to run code that restores any changes made to the application by the add-in and to perform general clean-up operations.
        /// An add-in can be unloaded in one of the following ways:
        /// - The user clears the check box next to the add-in in the COM Add-ins dialog box.
        /// - The host application closes. If the add-in is loaded when the application closes, it is unloaded. 
        ///   If the add-in's load behavior is set to Startup, it is reloaded when the application starts again.
        /// - The Connect property of the corresponding COMAddIn object is set to False.
        /// </summary>
        public event OnDisconnectionEventHandler OnDisconnection;

        /// <summary>
        /// The OnConnection event occurs when the COM add-in is loaded (connected). An add-in can be loaded in one of the following ways:
        /// The user starts the host application and the add-in's load behavior is specified to load when the application starts.
        /// The user loads the add-in in the COM Add-ins dialog box.
        /// The Connect property of the corresponding COMAddIn object is set to True.
        /// For more information about the COMAddIn object, search the Microsoft® Office Visual Basic Reference Help index for "COMAddIn object."
        /// </summary>
        public event OnConnectionEventHandler OnConnection;

        /// <summary>
        /// The OnAddInsUpdate event occurs when the set of loaded COM add-ins changes. 
        /// When an add-in is loaded or unloaded, the OnAddInsUpdate event occurs in any other loaded add-ins. 
        /// For example, if add-ins A and B both are loaded currently, and then add-in C is loaded, 
        /// the OnAddInsUpdate event occurs in add-ins A and B. If C is unloaded, the OnAddInsUpdate event occurs again in add-ins A and B. 
        /// </summary>
        public event OnAddInsUpdateEventHandler OnAddInsUpdate;

        /// <summary>
        /// The OnBeginShutdown event occurs when the host application begins its shutdown routines, 
        /// in the case where the application closes while the COM add-in is still loaded. 
        /// If the add-in is not loaded when the application closes, 
        /// the OnBeginShutdown event does not occur. When this event does occur, it occurs before the OnDisconnection event.
        /// You can use the OnBeginShutdown event procedure to run code when the user closes the application. For example, you can run code that saves form data to a file.
        /// </summary>
        public event OnBeginShutdownEventHandler OnBeginShutdown;

        private void RaiseOnStartupComplete(ref Array custom)
        {
            try
            {
                if (null != OnStartupComplete)
                    OnStartupComplete(ref custom);
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        private void RaiseShutdown(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != OnDisconnection)
                    OnDisconnection(RemoveMode, ref custom);
            }
            catch (System.Exception exception)
            {
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        private void RaiseOnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (null != OnConnection)
                    OnConnection(Application, ConnectMode, AddInInst, ref custom);
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        private void RaiseOnAddInsUpdate(ref Array custom)
        {
            try
            {
                if (null != OnAddInsUpdate)
                    OnAddInsUpdate(ref custom);
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        private void RaiseOnBeginShutdown(ref Array custom)
        {
            try
            {
                if (null != OnBeginShutdown)
                    OnBeginShutdown(ref custom);
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            RaiseOnStartupComplete(ref custom);
        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                this.Application = new Visio.Application(null, Application);
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            } 
            RaiseOnConnection(Application, ConnectMode, AddInInst, ref custom);
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            RaiseShutdown(RemoveMode, ref custom);
            try
            {
                if (!Application.IsDisposed)
                    Application.Dispose();
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseErrorHandlerMethod(exception, ref handled);
                if (!handled)
                    throw exception;
            } 
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            RaiseOnAddInsUpdate(ref custom);
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            RaiseOnBeginShutdown(ref custom);
        }

        #endregion

        #region ErrorHandler 
        
        /// <summary>
        /// Checks for a static method, signed with the ErrorHandlerAttribute and call them if its available
        /// </summary>
        /// <param name="type">type information for the class wtih static method </param>
        /// <param name="exception">occured exception</param>
        /// <param name="handled">must set to true when the error is handled by the client other the exception was thrown</param>
        private static void RaiseStaticErrorHandlerMethod(Type type, System.Exception exception, ref bool handled)
        {
            MethodInfo registerMethod = null;
            ErrorHandlerFunctionAttribute registerAttribute = null;
            bool methodPresent = AttributeHelper.GetErrorAttribute(type, ref registerMethod, ref registerAttribute);
            if (methodPresent)
                handled = (bool)registerMethod.Invoke(null, new object[] { exception });
        }

        /// <summary>
        /// checks for IErrorHandler implementation and call OnError, otherwhise redirect to RaiseStaticErrorHandlerMethod
        /// </summary>
        /// <param name="exception">occured exception</param>
        /// <param name="handled">must set to true when the error is handled by the client other the exception was thrown</param>
        private void RaiseErrorHandlerMethod(System.Exception exception, ref bool handled)
        {
            IErrorHandler handler = this as IErrorHandler;
            if (handler != null)
                handled = handler.OnError(exception);
            else
                RaiseStaticErrorHandlerMethod(Type, exception, ref handled);
        }
        
        #endregion

        #region COM Register Functions

        /// <summary>
        /// Called from regasm while register 
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComRegisterFunctionAttribute, Browsable(false), EditorBrowsable( EditorBrowsableState.Never)]
        public static void RegisterFunction(Type type)
        {
            try                
            {
                MethodInfo registerMethod = null;
                RegisterFunctionAttribute registerAttribute = null;
                bool registerMethodPresent = AttributeHelper.GetRegisterAttribute(type, ref registerMethod, ref registerAttribute);
                if (registerMethodPresent)
                {
                    CallDerivedRegisterMethod(type, registerMethod, registerAttribute);
                    if (registerAttribute.Value == RegisterMode.Replace)
                        return;
                }

                GuidAttribute guid = AttributeHelper.GetGuidAttribute(type);
                ProgIdAttribute progId = AttributeHelper.GetProgIDAttribute(type);
                RegistryLocationAttribute location = AttributeHelper.GetRegistryLocationAttribute(type);
				COMAddinAttribute addin = AttributeHelper.GetCOMAddinAttribute(type);

                Assembly thisAssembly = Assembly.GetAssembly(type);
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\1.0.0.0");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();
                
                key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32");
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

                // register addin in Visio
                Registry.CurrentUser.CreateSubKey(_addinOfficeRegistryKey +  progId.Value);
                RegistryKey regKeyVisio = null;
                
                if(location.Value == RegistrySaveLocation.LocalMachine)
                    regKeyVisio = Registry.LocalMachine.OpenSubKey(_addinOfficeRegistryKey + progId.Value, true);
                else
                    regKeyVisio = Registry.CurrentUser.OpenSubKey(_addinOfficeRegistryKey + progId.Value, true);

                regKeyVisio.SetValue("LoadBehavior", addin.LoadBehavior);
                regKeyVisio.SetValue("FriendlyName", addin.Name);
                regKeyVisio.SetValue("Description", addin.Description);
                if(-1 != addin.CommandLineSafe)
                    regKeyVisio.SetValue("CommandLineSafe", addin.CommandLineSafe);

                regKeyVisio.Close();

                 if( (registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallAfter))
                        registerMethod.Invoke(null, new object[] { type, RegisterCall.CallAfter });
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseStaticErrorHandlerMethod(type, exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        /// <summary>
        /// Called from regasm while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunctionAttribute, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            try
            {
                MethodInfo registerMethod = null;
                UnRegisterFunctionAttribute registerAttribute = null;
                bool registerMethodPresent = AttributeHelper.GetUnRegisterAttribute(type, ref registerMethod, ref registerAttribute);
                if (registerMethodPresent)
                {
                    CallDerivedUnRegisterMethod(type, registerMethod, registerAttribute);
                    if (registerAttribute.Value == RegisterMode.Replace)
                        return;
                }

                ProgIdAttribute progId = AttributeHelper.GetProgIDAttribute(type);
                RegistryLocationAttribute location = AttributeHelper.GetRegistryLocationAttribute(type);

                // unregister addin
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
               
                // unregister addin in office 
                if (location.Value == RegistrySaveLocation.LocalMachine)
                    Registry.LocalMachine.DeleteSubKey(_addinOfficeRegistryKey + progId.Value, false);
                else
                    Registry.CurrentUser.DeleteSubKey(_addinOfficeRegistryKey + progId.Value, false);

                if ((registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallAfter))
                    registerMethod.Invoke(null, new object[] { type, RegisterCall.CallAfter });
            }
            catch (System.Exception exception)
            {
				NetOffice.DebugConsole.WriteException(exception);
                bool handled = false;
                RaiseStaticErrorHandlerMethod(type, exception, ref handled);
                if (!handled)
                    throw exception;
            }
        }

        /// <summary>
        /// Derived Register Call Helper
        /// </summary>
        /// <param name="type">type for derived class</param>
        /// <param name="registerMethod">the method to call</param>
        /// <param name="registerAttribute">arguments</param>
        private static void CallDerivedRegisterMethod(Type type, MethodInfo registerMethod, RegisterFunctionAttribute registerAttribute)
        {
            if (registerAttribute.Value == RegisterMode.Replace)
                registerMethod.Invoke(null, new object[] { type, RegisterCall.Replace });
            else if (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallBefore)
                registerMethod.Invoke(null, new object[] { type, RegisterCall.CallBefore });
        }

        /// <summary>
        /// Derived Unregister Call Helper
        /// </summary>
        /// <param name="type">type for derived class</param>
        /// <param name="registerMethod">the method to call</param>
        /// <param name="registerAttribute">arguments</param>
        private static void CallDerivedUnRegisterMethod(Type type, MethodInfo registerMethod, UnRegisterFunctionAttribute registerAttribute)
        {
            if (registerAttribute.Value == RegisterMode.Replace)
                registerMethod.Invoke(null, new object[] { type, RegisterCall.Replace });
            else if (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallBefore)
                registerMethod.Invoke(null, new object[] { type, RegisterCall.CallBefore });
        }

        #endregion

        #region Private Helper Methods

        /// <summary>
        /// reads text file from ressource
        /// </summary>
        /// <param name="fileName">ressourceLocation</param>
        /// <returns>text content</returns>
        private string ReadRessourceFile(string fileName)
        {
            Assembly assembly = Type.Assembly;
            System.IO.Stream ressourceStream = assembly.GetManifestResourceStream(fileName);
            if (ressourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(ressourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource File."));

            string text = textStreamReader.ReadToEnd();
            ressourceStream.Close();
            textStreamReader.Close();
            return text;
        }

        #endregion
    }
}