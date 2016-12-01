using System;
using NetRuntimeSystem = System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;
using PowerPoint = NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Tools
{
    /// <summary>
    /// NetOffice MS-PowerPoint COM Addin
    /// </summary>
	[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddin : COMAddinBase, IOfficeCOMAddin
    {
        #region Fields

        /// <summary>
        /// MS-PowerPoint Registry Path 
        /// </summary>
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\PowerPoint\\Addins\\";

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
            TaskPanes = new CustomTaskPaneCollection();
			TaskPaneInstances = new List<ITaskPane>();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Common Tasks Helper. The property is available after the host application has called OnConnection for the instance
        /// </summary>
        public Utils.CommonUtils Utils { get; private set; }     

        /// <summary>
        /// Host Application Instance
        /// </summary>
        protected internal PowerPoint.Application Application { get; private set; }
        
        /// <summary>
        /// Collection with all created custom Task Panes
        /// </summary>
        protected CustomTaskPaneCollection TaskPanes { get; private set; }

        /// <summary>
        /// TaskPaneFactory from CTPFactoryAvailable
        /// </summary>
        protected Office.ICTPFactory TaskPaneFactory { get; set; }

		/// <summary>
        /// ITaskPane Instances
        /// </summary>
		protected List<ITaskPane> TaskPaneInstances { get; set; }

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
            if (null != TaskPaneFactory)
                result.Add(TaskPaneFactory);

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
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
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
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
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
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnConnection, exception);
            }
        }

        private void RaiseOnAddInsUpdate(ref Array custom)
        {
            try
            {
                if (null != OnAddInsUpdate)
                    OnAddInsUpdate(ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
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
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                Tweaks.ApplyTweaks(Factory, this, Type, "PowerPoint", IsLoadedFromSystem);
                LoadingTimeElapsed = (DateTime.Now - _creationTime);
                Roots = OnCreateRoots();
                RaiseOnStartupComplete(ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        void IDTExtensibility2.OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (null != custom && custom.Length > 0)
                {
                    object firstCustomItem = custom.GetValue(1);
                    string tryString = null != firstCustomItem ? firstCustomItem.ToString() : String.Empty;
                    NetRuntimeSystem.Int32.TryParse(tryString, out _automationCode);                    
                }

                this.Application = new PowerPoint.Application(Factory, null, application);
                Utils = OnCreateUtils();
                RaiseOnConnection(this.Application, ConnectMode, AddInInst, ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnConnection, exception);
            }
        }

        void IDTExtensibility2.OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try
            {
                foreach (ITaskPane item in TaskPaneInstances)
                {
                    try
                    {
                        item.OnDisconnection();
                    }
                    catch (NetRuntimeSystem.Exception exception)
                    {
                        NetOffice.DebugConsole.Default.WriteException(exception);
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
                        NetOffice.DebugConsole.Default.WriteException(exception);
                    }
                }

                try
                {
                    if (null != TaskPaneFactory && false == TaskPaneFactory.IsDisposed)
                        TaskPaneFactory.Dispose();
                }
                catch (NetRuntimeSystem.Exception exception)
                {
                    NetOffice.DebugConsole.Default.WriteException(exception);
                }

                try
                {
                    Tweaks.DisposeTweaks(Factory, this, Type);
                    RaiseOnDisconnection(RemoveMode, ref custom);
                    Utils.Dispose();

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
                    NetOffice.DebugConsole.Default.WriteException(exception);
                }	
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
            }
        }

        void IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            try
            {
                RaiseOnAddInsUpdate(ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnAddInsUpdate, exception);
            }
        }

        void IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            try
            {
                RaiseOnBeginShutdown(ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        #endregion

        #region IRibbonExtensibility Members

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="RibbonID">target ribbon id, only used from Outlook and ignored in this standard implementation. overwrite this method if you need a custom behavior</param>
        /// <returns>XML content or String.Empty</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual string GetCustomUI(string RibbonID)
        {
            try
            {
                CustomUIAttribute ribbon = AttributeReflector.GetRibbonAttribute(Type);
                if (null != ribbon)
                    return Utils.Resource.ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
                else
                    return String.Empty;
            }
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
                return String.Empty;
            } 
        }

        #endregion

        #region ICustomTaskPaneConsumer Member

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="CTPFactoryInst">factory proxy from host application</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual void CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
                if (null == CTPFactoryInst)
                {
                    Factory.Console.WriteLine("Warning: null argument recieved in CTPFactoryAvailable. argument name: CTPFactoryInst");
                    return;
                }

                CustomTaskPaneHandler paneHandler = new CustomTaskPaneHandler();
                paneHandler.ProceedCustomPaneAttributes(TaskPanes, Type, CallOnCreateTaskPaneInfo, AttributePane_VisibleStateChange, AttributePane_DockPositionStateChange);
                TaskPaneFactory = paneHandler.CreateCustomPanes<ITaskPane, PowerPoint.Application>(Factory, CTPFactoryInst, TaskPanes, TaskPaneInstances, OnError, Application);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        /// <summary>
        /// The method is called while the CustomPane attribute is processed
        /// </summary>
        /// <param name="paneInfo">pane definition</param>
		/// <returns>true if paneInfo is modified, otherwise false to set the default or attribute values</returns>
		protected internal virtual bool OnCreateTaskPaneInfo(TaskPaneInfo paneInfo)
		{
			return false;
		}
		
        /// <summary>
        /// Called after any visibility changes
        /// </summary>
        /// <param name="customTaskPaneInst">pane instance</param>
		protected internal virtual void TaskPaneVisibleStateChanged(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
		{
	
		}

		/// <summary>
        /// Called after any position changes but not for size changes
        /// </summary>
        /// <param name="customTaskPaneInst">pane instance</param>
		protected internal virtual void TaskPaneDockStateChanged(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
		{
			
		}

		private void CallTaskPaneVisibleStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
		{
			try
			{
				foreach(TaskPaneInfo item in TaskPanes)
				{
					if(item.Pane == customTaskPaneInst)
					{
						try
						{
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
							if (null != target && item.Pane == customTaskPaneInst)
							{
								try
                                {
									target.OnVisibleStateChanged(item.Pane.Visible);
								}
								catch(Exception exception)
								{
									Factory.Console.WriteException(exception);
								}
							}
						}
						catch(Exception exception)
						{
							Factory.Console.WriteException(exception);
						}
					}
				}
                TaskPaneVisibleStateChanged(customTaskPaneInst);
			}
			catch(Exception exception)
			{
			   Factory.Console.WriteException(exception);
			}
		}

		private void CallTaskPaneDockPositionStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
		{
			try
			{
				foreach(TaskPaneInfo item in TaskPanes)
				{
					if(item.Pane == customTaskPaneInst)
					{
						try
						{
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
							if (null != target && item.Pane == customTaskPaneInst)
							{
								try
								{
                                    target.OnDockPositionChanged(item.Pane.DockPosition);
								}
								catch(Exception exception)
								{
									Factory.Console.WriteException(exception);
								}
							}
						}
						catch(Exception exception)
						{
							Factory.Console.WriteException(exception);
						}
					}
				}
                TaskPaneDockStateChanged(customTaskPaneInst);
			}
			catch(Exception exception)
			{
			   Factory.Console.WriteException(exception);
			}
		}

		private bool CallOnCreateTaskPaneInfo(TaskPaneInfo paneInfo)
		{
			try
			{
				return OnCreateTaskPaneInfo(paneInfo);
			}
			catch(Exception exception)
			{
				Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CTPFactoryAvailable, exception);
				return false;
			}
		}
		
        private void AttributePane_VisibleStateChange(NetOffice.OfficeApi._CustomTaskPane CustomTaskPaneInst)
        {           
			try
			{
				CallTaskPaneVisibleStateChange(CustomTaskPaneInst);
			}
			catch(Exception exception)
			{
				Factory.Console.WriteException(exception);
			}
        }

        private void AttributePane_DockPositionStateChange(Office._CustomTaskPane CustomTaskPaneInst)
        {
			try
			{
                CallTaskPaneDockPositionStateChange(CustomTaskPaneInst);
			}
			catch(Exception exception)
			{
				Factory.Console.WriteException(exception);
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
        /// You have no warranties for dispose your tweak.
        /// </summary>
        /// <param name="name">name for the tweak</param>
        /// <param name="value">value for the teak</param>
        protected virtual void DisposeCustomTweak(string name, string value)
        {

        }

        /// <summary>
        /// Creates an registry tweak entry in the current addin key
        /// </summary>
        /// <param name="addinType">addin type information</param>
        /// <param name="name">name for the tweak</param>
        /// <param name="value">value for the tweak</param>
        /// <param name="throwException">throw exception on error</param>
        /// <returns>true if key was created otherwise false</returns>
        protected static bool SetTweakPersistenceEntry(Type addinType, string name, string value, bool throwException)
        {

            return OfficeApi.Tools.COMAddin.SetTweakPersistenceEntry(
                                                          ApplicationIdentifiers.ApplicationType.PowerPoint,
                                                          addinType,
                                                          name, value,
                                                          throwException);
        }

        #endregion

        #region Virtual Methods

        /// <summary>
        /// Create the used utils. The method was called in OnConnection
        /// </summary>
        /// <returns>new ToolsUtils instance</returns>
        protected internal virtual Utils.CommonUtils OnCreateUtils()
        {
            return new Utils.CommonUtils(this, Type, 3 == _automationCode ? true : false, this.Type.Assembly);
        }

        /// <summary>
        /// Create the used factory. The method was called as first in the base ctor
        /// </summary>
        /// <returns>new Settings instance</returns>
        protected virtual Core CreateFactory()
        {
            Core core = new Core();
            ForceInitializeAttribute attribute = AttributeReflector.GetForceInitializeAttribute(Type);
            if (null != attribute)
            {
                core.Settings.EnableDebugOutput = attribute.EnableDebugOutput;
                core.CheckInitialize();
            }
            return core;
        }

        /// <summary>
        /// Create the necessary factory and was called in the first line in base ctor
        /// </summary>
        /// <returns></returns>
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

        #region Private Methods

        /// <summary>
        /// Try to detect the addin is loaded from system hive key
        /// </summary>
        /// <returns>null if unkown or true/false</returns>
        private bool? IsLoadedFromSystem()
        {
            if (null != _isLoadedFromSystem)
                return _isLoadedFromSystem;

            OfficeApi.Tools.Utils.RegistryLocationResult result = OfficeApi.Tools.Utils.CommonUtils.TryFindAddinLoadLocation(Type,
                    ApplicationIdentifiers.ApplicationType.PowerPoint);
            switch (result)
            {
                case Office.Tools.Utils.RegistryLocationResult.User:
                    _isLoadedFromSystem = false;
                    break;
                case Office.Tools.Utils.RegistryLocationResult.System:
                    _isLoadedFromSystem = true;
                    break;
                default:
                    throw new IndexOutOfRangeException();
            }

            return _isLoadedFromSystem;
        }

        #endregion

        #region ErrorHandler 

        /// <summary>
        /// Custom error handler
        /// </summary>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnError(ErrorMethodKind methodKind, NetRuntimeSystem.Exception exception)
        {

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
            RegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeRegisterKeyState.NeedToCreate);
        }

        /// <summary>
        /// Called from regasm while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunctionAttribute, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            UnRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeUnRegisterKeyState.NeedToDelete);
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
            InstallScope currentScope = (InstallScope)scope;
            OfficeRegisterKeyState currentKeyState = (OfficeRegisterKeyState)keyState;

            RegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, currentScope, currentKeyState);
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
            InstallScope currentScope = (InstallScope)scope;
            OfficeUnRegisterKeyState currentKeyState = (OfficeUnRegisterKeyState)keyState;

            UnRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, currentScope, currentKeyState);
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