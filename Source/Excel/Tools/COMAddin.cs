using System;
using NetRuntimeSystem = System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;
using Excel = NetOffice.ExcelApi;
using NetOffice.OfficeApi.Enums;

namespace NetOffice.ExcelApi.Tools
{
    /// <summary>
    /// NetOffice MS-Excel COM Addin
    /// </summary>
    public abstract class COMAddin : OfficeCOMAddin, IOfficeCOMAddin
    {
        #region Fields

        /// <summary>
        /// MS-Excel Registry Path
        /// </summary>
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\Excel\\Addins\\";

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
        /// Common Tasks Helper. The property is available after the host application has called OnConnection for the instance
        /// </summary>
        public Contribution.CommonUtils Utils { get; private set; }

        /// <summary>
        /// Host Application Instance
        /// </summary>
        public new Excel.Application Application
        {
            get
            {
                return base.Application as Excel.Application;
            }
        }

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

        #endregion

        #region IDTExtensibility2 Overrides

        /// <summary>
        /// Occurs whenever an add-in is loaded into MS-Office
        /// </summary>
        /// <param name="application">A reference to an instance of the office application</param>
        /// <param name="connectMode">An ext_ConnectMode enumeration value that indicates the way the add-in was loaded into MS-Office</param>
        /// <param name="addInInst">An AddIn reference to the add-in's own instance. This is stored for later use, such as determining the parent collection for the add-in</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected override void HandleOnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            if (null != custom && custom.Length > 0)
            {
                object firstCustomItem = custom.GetValue(1);
                string tryString = null != firstCustomItem ? firstCustomItem.ToString() : String.Empty;
                System.Int32.TryParse(tryString, out _automationCode);
            }

            base.Application = Factory.CreateKnownObjectFromComProxy<Excel.Application>(null, application, typeof(Excel.Application));
            Utils = OnCreateUtils();
            TryCreateCustomObject(addInInst);
            RaiseOnConnection(this.Application, connectMode, addInInst, ref custom);
        }

        /// <summary>
        /// Occurs whenever an add-in is unloaded from MS Office
        /// </summary>
        /// <param name="removeMode">An ext_DisconnectMode enumeration value that informs an add-in why it was unloaded.</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use after the add-in unloads</param>
        protected override void HandleOnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                RaiseOnDisconnection(removeMode, ref custom);
                Utils.Dispose();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
            }

            foreach (ITaskPane item in TaskPaneInstances)
            {
                try
                {
                    item.OnDisconnection();
                }
                catch (NetRuntimeSystem.Exception exception)
                {
                    Factory.Console.WriteException(exception);
                }
            }

            foreach (var item in TaskPanes)
            {
                try
                {
                    if (null != item.Pane && !item.Pane.IsDisposed)
                        item.Pane.Dispose();
                }
                catch (NetRuntimeSystem.Exception exception)
                {
                    Factory.Console.WriteException(exception);
                }
            }

            try
            {
                if (null != TaskPaneFactory && false == TaskPaneFactory.IsDisposed)
                    TaskPaneFactory.Dispose();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
            }

            try
            {
                if (null != RibbonUI)
                {
                    RibbonUI.Dispose();
                    RibbonUI = null;
                }
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
            }

            try
            {
                if (!Application.IsDisposed)
                    Application.Dispose();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
            }

            try
            {
                CleanUp();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        /// <summary>
        ///  Occurs whenever an add-in, which is set to load when MS Office starts, loads.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use when the add-in loads</param>
        protected override void HandleOnStartupComplete(ref Array custom)
        {
            LoadingTimeElapsed = (DateTime.Now - _creationTime);
            Roots = OnCreateRoots();
            RaiseOnStartupComplete(ref custom);
        }

        /// <summary>
        /// Occurs whenever an add-in is loaded or unloaded from MS Office
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected override void HandleOnAddInsUpdate(ref Array custom)
        {
            RaiseOnAddInsUpdate(ref custom);
        }

        /// <summary>
        /// Occurs whenever MS Office shuts down while an add-in is running
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected override void HandleOnBeginShutdown(ref Array custom)
        {
            RaiseOnBeginShutdown(ref custom);
        }

        #endregion

        #region Virtual Methods

        /// <summary>
        /// Create the used utils. The method was called in OnConnection
        /// </summary>
        /// <returns>new ToolsUtils instance</returns>
        protected internal virtual Contribution.CommonUtils OnCreateUtils()
        {
            return new Contribution.CommonUtils(this, Type, 3 == _automationCode ? true : false, this.Type.Assembly);
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
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.CreateFactory, exception);
                return null;
            }
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

            COMAddinRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeRegisterKeyState.NeedToCreate);
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

            COMAddinUnRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeUnRegisterKeyState.NeedToDelete);
        }

        /// <summary>
        /// Called from RegAddin while register
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">NetOffice.Tools.InstallScope enum value</param>
        /// <param name="keyState">NetOffice.Tools.OfficeRegisterKeyState enum value</param>
        [ComRegisterCall]
        private static void OptimizedRegisterFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            InstallScope currentScope = (InstallScope)scope;
            OfficeRegisterKeyState currentKeyState = (OfficeRegisterKeyState)keyState;

            COMAddinRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, currentScope, currentKeyState);
        }

        /// <summary>
        /// Called from RegAddin while unregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">NetOffice.Tools.InstallScope enum value</param>
        /// <param name="keyState">NetOffice.Tools.OfficeUnRegisterKeyState enum value</param>
        [ComUnregisterCall]
        private static void OptimizedUnregisterFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            InstallScope currentScope = (InstallScope)scope;
            OfficeUnRegisterKeyState currentKeyState = (OfficeUnRegisterKeyState)keyState;

            COMAddinUnRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, currentScope, currentKeyState);
        }

        /// <summary>
        /// Called from RegAddin while export registry informations
        /// </summary>
        /// <param name="type">Type information for the class</param>
        /// <param name="scope">NetOffice.Tools.InstallScope enum value</param>
        /// <param name="keyState">NetOffice.Tools.OfficeRegisterKeyState enum value</param>
        /// <returns>Registry keys/values to be add in the registry export or null</returns>
        [ComRegExportCall]
        private static RegExport RegExportFunction(Type type, int scope, int keyState)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            InstallScope currentScope = (InstallScope)scope;
            OfficeRegisterKeyState currentKeyState = (OfficeRegisterKeyState)keyState;

            return RegExportHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, currentScope, currentKeyState);
        }

        #endregion
    }
}
