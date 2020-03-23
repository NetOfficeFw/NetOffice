﻿using System;
using System.Runtime.InteropServices;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.Tools;
using NetOffice.VBIDEApi;

namespace NetOffice.VBIDEApi.Tools
{
    /// <summary>
    /// NetOffice VBIDE COM Addin
    /// Notes: HKEY_LOCAL_MACHINE hive is ignored by the VBA editor
    /// </summary>
    [ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddin : COMAddinBase, ICOMAddin
    {
        #region Fields

        /// <summary>
        /// VBIDE Registry Path 
        /// </summary>
        private static readonly string _addinOfficeRegistryKey = "Software\\Microsoft\\VBA\\VBE\\6.0\\Addins\\";
       
        /// <summary>
        /// VBIDE Registry Path 64 Bit
        /// </summary>
        private static readonly string _addinOfficeRegistryKey64 = "Software\\Microsoft\\VBA\\VBE\\6.0\\Addins64\\";

        /// <summary>
        /// First field in OnConnection custom argument array
        /// </summary>
        private int _automationCode = -1;

        /// <summary>
        /// Cache field used in IsLoadedFromSystem() method
        /// </summary>
        private bool? _isLoadedFromSystem;

        /// <summary>
        /// Instance factory to avoid trouble with addins in same appdomain
        /// </summary>
        private Core _factory;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public COMAddin()
        {
            _factory = RaiseCreateFactory();
            if (null == _factory)
                _factory = Core.Default;        
        }

        #endregion

        #region Properties

        /// <summary>
        /// Host Application Instance
        /// </summary>
        protected internal VBE Application { get; private set; }

        /// <summary>
        /// Custom addin object if created
        /// </summary>
        protected internal object CustomObject { get; private set; }

        /// <summary>
        /// Cached Error Method Delegate
        /// </summary>
        private MethodInfo ErrorMethod { get; set; }

        /// <summary>
        /// Cached Register Error Method Delegate
        /// </summary>
        private static MethodInfo RegisterErrorMethod { get; set; }

        #endregion

        #region COMAddinBase

        /// <summary>
        /// Generic Host Application Instance
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public override ICOMObject AppInstance
        {
            get { return Application; }
        }

        /// <summary>
        /// The used factory core
        /// </summary>
        public override Core Factory
        {
            get
            {
                return _factory;
            }
        }
       
        /// <summary>
        /// Instance managed root com objects
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public override IEnumerable Roots { get; protected set; }

        /// <summary>
        /// Returns an enumerable sequence with instance managed com objects on root level
        /// </summary>
        /// <returns>ICOMObject enumerator</returns>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        protected internal virtual IEnumerable<ICOMObject> OnCreateRoots()
        {
            List<ICOMObject> result = new List<ICOMObject>();
            result.Add(Application);

            return result.ToArray();
        }

        #endregion

        #region IDTExtensibility2 Events

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
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        private void RaiseOnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                if (null != OnDisconnection)
                    OnDisconnection(RemoveMode, ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
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
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
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
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnAddInsUpdate, exception);
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
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        #endregion

        #region IDTExtensibility2 Members

        void NetOffice.Tools.Native.IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {             
                LoadingTimeElapsed = (DateTime.Now - _creationTime);
                Roots = OnCreateRoots();
                RaiseOnStartupComplete(ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (null != custom && custom.Length > 0)
                {
                    object firstCustomItem = custom.GetValue(1);
                    string tryString = null != firstCustomItem ? firstCustomItem.ToString() : String.Empty;
                    System.Int32.TryParse(tryString, out _automationCode);
                }

                this.Application = new VBE(null, application);
                TryCreateCustomObject(AddInInst);
                RaiseOnConnection(this.Application, ConnectMode, AddInInst, ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnConnection, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {                 
                try
                {                  
                    RaiseOnDisconnection(RemoveMode, ref custom);
                    Tweaks.DisposeTweaks(Factory, this, Type);
                }
                catch (System.Exception exception)
                {
                    Factory.Console.WriteException(exception);
                }
                
                try
                {
                    if (!Application.IsDisposed)
                        Application.Dispose();
                }
                catch (System.Exception exception)
                {
                    Factory.Console.WriteException(exception);
                }
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            try
            {
                RaiseOnAddInsUpdate(ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnAddInsUpdate, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            try
            {
                RaiseOnBeginShutdown(ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        #endregion

        #region Tweaks

        /// <summary>
        /// This is method is called while startup and ask for permissions to apply a tweak. 
        /// </summary>
        /// <param name="name">name of the tweak</param>
        /// <param name="value">value of the tweak</param>
        /// <returns>true(default) or false if you dont want this tweak is affected to the addin instance</returns>
        protected virtual bool AllowApplyTweak(string name, string value)
        {
            return true;
        }

        /// <summary>
        /// Called for custom tweaks to apply the tweak.
        /// </summary>
        /// <param name="name">name for the tweak</param>
        /// <param name="value">value for the teak</param>
        protected virtual void ApplyCustomTweak(string name, string value)
        {
        }

        /// <summary>
        /// Called for custom tweaks to unload a tweak. Please note: This method is not called in case of unexpected termination.
        /// You have no warranties for disposing your tweak.
        /// </summary>
        /// <param name="name">name for the tweak</param>
        /// <param name="value">value for the teak</param>
        protected virtual void DisposeCustomTweak(string name, string value)
        {

        }

        #endregion

        #region Methods

        /// <summary>
        /// Returns an instance to publish them as addin custom object.
        /// External code like vba can access this object if instance is available as COM component.
        /// This object is available as Appplication.COMAddins(?).Object
        /// </summary>
        /// <returns>addin instance object or null(Nothing in Visual Basic)</returns>
        protected virtual object OnCreateObjectInstance()
        {
            return null;
        }

        /// <summary>
        /// Try to create a custom addin object instance
        /// </summary>
        /// <param name="addInInst">given instance from OnConnection event</param>
        private void TryCreateCustomObject(object addInInst)
        {
            try
            {
                CustomObject = OnCreateObjectInstance();
                if (null != CustomObject)
                {
                    object[] param = new object[1];
                    param[0] = CustomObject;
                    addInInst.GetType().InvokeMember("Object", System.Reflection.BindingFlags.SetProperty, null, addInInst, param);
                }
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CreateCustomAddinInstance, exception);
            }
        }

        /// <summary>
        /// Create the used factory. The method was called as first in the base ctor
        /// </summary>
        /// <returns>new Core instance</returns>
        protected internal virtual Core CreateFactory()
        {
            Core core = new Core();
            ForceInitializeAttribute attribute = AttributeReflector.GetForceInitializeAttribute(Type);
            if (null != attribute)
            {
                core.Settings.EnableMoreDebugOutput = attribute.EnableMoreDebugOutput;
                core.CheckInitialize();
            }
            return core;
        }

        /// <summary>
        /// Create the necessary factory and was called in the first line in base ctor
        /// </summary>
        /// <returns>new Core instance or null if failed</returns>
        private Core RaiseCreateFactory()
        {
            try
            {
                return CreateFactory();
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.CreateFactory, exception);
                return null;
            }
        }

        #endregion

        #region ErrorHandler

        /// <summary>
        /// Custom error handler
        /// </summary>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnError(ErrorMethodKind methodKind, System.Exception exception)
        {

        }

        #endregion

        #region COM Register Functions

        /// <summary>
        /// Called from RegAsm while register 
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComRegisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void RegisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            COMAddinRegisterHandler.ProceedUser(type, new string[] { _addinOfficeRegistryKey, _addinOfficeRegistryKey64 }, OfficeRegisterKeyState.NeedToCreate);
        }

        /// <summary>
        /// Called from RegAsm while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunction, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            COMAddinUnRegisterHandler.ProceedUser(type, new string[] { _addinOfficeRegistryKey, _addinOfficeRegistryKey64 }, OfficeUnRegisterKeyState.NeedToDelete);
        }

        /// <summary>
        /// Called from RegAddin while register
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">(ignored)</param>
        /// <param name="keyState">NetOffice.Tools.OfficeRegisterKeyState enum value</param>
        [ComRegisterCall]
        private static void OptimizedRegisterFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            OfficeRegisterKeyState currentKeyState = (OfficeRegisterKeyState)keyState;
            COMAddinRegisterHandler.ProceedUser(type, new string[] { _addinOfficeRegistryKey, _addinOfficeRegistryKey64 }, currentKeyState);
        }

        /// <summary>
        /// Called from RegAddin while unregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">(ignored)</param>
        /// <param name="keyState">NetOffice.Tools.OfficeUnRegisterKeyState enum value</param>
        [ComUnregisterCall]
        private static void OptimizedUnregisterFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            OfficeUnRegisterKeyState currentKeyState = (OfficeUnRegisterKeyState)keyState;
            COMAddinUnRegisterHandler.ProceedUser(type, new string[] { _addinOfficeRegistryKey, _addinOfficeRegistryKey64 }, currentKeyState);
        }

        /// <summary>
        /// Called from RegAddin while export registry information 
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">(ignored)</param>
        /// <param name="keyState">NetOffice.Tools.OfficeRegisterKeyState enum value</param>
        /// <returns>Registry keys/values to be add in the registry export or null</returns>
        [ComRegExportCall]
        private static RegExport RegExportFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
           
            OfficeRegisterKeyState currentKeyState = (OfficeRegisterKeyState)keyState;
            return RegExportHandler.ProceedUser(type, new string[] { _addinOfficeRegistryKey, _addinOfficeRegistryKey64 }, currentKeyState);
        }

        #endregion
    }
}