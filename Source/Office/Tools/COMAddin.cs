using System;
using NetRuntimeSystem = System;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// NetOffice COM Addin
    /// </summary>
	[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddin : COMAddinBase, IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        #region Fields

        /// <summary>
        /// MS-Office Registry Path 
        /// </summary>
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\{0}\\AddIns\\";

        /// <summary>
        /// First field in OnConnection custom argument array
        /// </summary>
        private int _automationCode = -1;

        #endregion
        
        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public COMAddin()
        {
            Factory = RaiseCreateFactory();
            if (null == Factory)
                Factory = Core.Default;
            TaskPanes = new CustomTaskPaneCollection();
			TaskPaneInstances = new List<ITaskPane>();
            Type = this.GetType();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Common Tasks Helper. The property is available after the host application has called OnConnection for the instance
        /// </summary>
        public CommonUtils Utils { get; private set; }

        /// <summary>
        /// The used factory core
        /// </summary>
        public Core Factory { get; private set; }

        /// <summary>
        /// Type Information of the instance
        /// </summary>
        protected Type Type { get; set; }

        /// <summary>
        /// Host Application Instance
        /// </summary>
        protected internal COMObject Application { get; private set; }
        
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
            catch (Exception exception)
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
            catch (Exception exception)
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
            catch (Exception exception)
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
            catch (Exception exception)
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
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        #endregion

        #region COMAddinBase

        /// <summary>
        /// Generic Host Application Instance
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public override COMObject AppInstance
        {
            get { return Application; }
        }

        #endregion

        #region IDTExtensibility2 Members

        void IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                RaiseOnStartupComplete(ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        void IDTExtensibility2.OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (AttributeHelper.GetTweakAttribute(Type).Enabled == true)
                {
                    string registryEndPoint = TryDetectHostRegistryKey(Application);
                    if (null != registryEndPoint)
                        Tweaks.ApplyTweaks(Factory, this, Type, registryEndPoint);
                }

                if (custom.Length > 0)
                {
                    object firstCustomItem = custom.GetValue(1);
                    string tryString = null != firstCustomItem ? firstCustomItem.ToString() : String.Empty;
                    System.Int32.TryParse(tryString, out _automationCode);
                }

                this.Application = Factory.CreateObjectFromComProxy(null, Application);
                Utils = OnCreateUtils();
                RaiseOnConnection(Application, ConnectMode, AddInInst, ref custom);
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
                    catch (System.Exception exception)
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
                    catch (System.Exception exception)
                    {
                        NetOffice.DebugConsole.Default.WriteException(exception);
                    }
                }

                try
                {
                    if (null != TaskPaneFactory && false == TaskPaneFactory.IsDisposed)
                        TaskPaneFactory.Dispose();
                }
                catch (System.Exception exception)
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
                catch (System.Exception exception)
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
        public virtual string GetCustomUI(string RibbonID)
        {
            try
            {
                
                CustomUIAttribute ribbon = AttributeHelper.GetRibbonAttribute(Type);
                if (null != ribbon)
                    return Utils.Resource.ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
                else
                    return string.Empty;
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
				return string.Empty;
            } 
        }

        #endregion

        #region ICustomTaskPaneConsumer Member

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="CTPFactoryInst">factory proxy from host application</param>
        public virtual void CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
                if (null == CTPFactoryInst)
                {
                    Factory.Console.WriteLine("Warning: null argument recieved in CTPFactoryAvailable. argument name: CTPFactoryInst");
                    return;
                }

				CustomPaneAttribute paneAttribute = AttributeHelper.GetCustomPaneAttribute(Type);
				if(null != paneAttribute)
				{
					TaskPaneInfo item = TaskPanes.Add(paneAttribute.PaneType, paneAttribute.PaneType.Name);
					if(!CallOnCreateTaskPaneInfo(item))
					{
						item.Title = paneAttribute.Title;
						item.Visible = paneAttribute.Visible;
                        item.DockPosition = (Office.Enums.MsoCTPDockPosition)Enum.Parse(typeof(Office.Enums.MsoCTPDockPosition), paneAttribute.DockPosition.ToString());
                        item.DockPositionRestrict = (Office.Enums.MsoCTPDockPositionRestrict)Enum.Parse(typeof(Office.Enums.MsoCTPDockPositionRestrict), paneAttribute.DockPositionRestrict.ToString());
                        item.Width = paneAttribute.Width;
                        item.Height = paneAttribute.Height;
						item.Arguments = new object[] { this };
					}

					item.VisibleStateChange += new NetOffice.OfficeApi.CustomTaskPane_VisibleStateChangeEventHandler(AttributePane_VisibleStateChange);
					item.DockPositionStateChange += new Office.CustomTaskPane_DockPositionStateChangeEventHandler(AttributePane_DockPositionStateChange);
				}

                TaskPaneFactory = new NetOffice.OfficeApi.ICTPFactory(Factory, null, CTPFactoryInst);
                foreach (TaskPaneInfo item in TaskPanes)
                {
                    string title = item.Title;
                    Office.CustomTaskPane taskPane = TaskPaneFactory.CreateCTP(item.Type.FullName, title) as Office.CustomTaskPane;
                    item.Pane = taskPane;
                    item.AssignEvents();
                    item.IsLoaded = true;

                    switch (taskPane.DockPosition)
                    {
                        case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionLeft:
                        case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionRight:
                            taskPane.Width = item.Width;
                            break;
                        case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionTop:
                        case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionBottom:
                            taskPane.Height = item.Height;
                            break;
                        case NetOffice.OfficeApi.Enums.MsoCTPDockPosition.msoCTPDockPositionFloating:
                            item.Width = paneAttribute.Width;
                            taskPane.Height = item.Height;
                            break;
                        default:
                            break;
                    }

                    ITaskPane pane = taskPane.ContentControl as ITaskPane;
                    if (null != pane)
                    {
                        TaskPaneInstances.Add(pane);
                        object[] argumentArray = new object[0];

                        if (item.Arguments != null)
                            argumentArray = item.Arguments;

                        pane.OnConnection(Application, taskPane, argumentArray);
                    }

                    foreach (KeyValuePair<string, object> property in item.ChangedProperties)
                    {
                        if (property.Key == "Title")
                            continue;

                        try
                        {
                            if (property.Key == "Width") // avoid to set width in top and bottom align
                            {
                                object outValue = null;
                                item.ChangedProperties.TryGetValue("DockPosition", out outValue);
                                if (null != outValue)
                                {

                                    Office.Enums.MsoCTPDockPosition position = (Office.Enums.MsoCTPDockPosition)Enum.Parse(typeof(Office.Enums.MsoCTPDockPosition), outValue.ToString());
                                    if (position == Office.Enums.MsoCTPDockPosition.msoCTPDockPositionTop || position == Office.Enums.MsoCTPDockPosition.msoCTPDockPositionBottom)
                                        continue;
                                }
                            }

                            if (property.Key == "Height")   // avoid to set height in left and right align
                            {
                                object outValue = null;
                                item.ChangedProperties.TryGetValue("DockPosition", out outValue);
                                if (null == outValue)
                                    outValue = Office.Enums.MsoCTPDockPosition.msoCTPDockPositionRight; // NetOffice default position if unset

                                Office.Enums.MsoCTPDockPosition position = (Office.Enums.MsoCTPDockPosition)Enum.Parse(typeof(Office.Enums.MsoCTPDockPosition), outValue.ToString());
                                if (position == Office.Enums.MsoCTPDockPosition.msoCTPDockPositionLeft || position == Office.Enums.MsoCTPDockPosition.msoCTPDockPositionRight)
                                    continue;
                            }

                            taskPane.GetType().InvokeMember(property.Key, BindingFlags.SetProperty, null, taskPane, new object[] { property.Value });
                        }
                        catch
                        {
                            ;
                        }
                    }
                }
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
            try
            {
                if (null == addinType)
                    return false;
                RegistryLocationAttribute registry = AttributeHelper.GetRegistryLocationAttribute(addinType);
                ProgIdAttribute progID = AttributeHelper.GetProgIDAttribute(addinType);
                MultiRegisterAttribute register = MultiRegisterAttribute.GetAttribute(addinType);

                if (null == registry)
                    return false;
                if (null == progID)
                    return false;
                // my current keyboard miss the logical or. thanks LogiLink

                foreach (RegisterIn item in register.Products)
                {
                    RegistryKey regKeyOffice = null;
                    if (registry.Value == RegistrySaveLocation.LocalMachine)
                        regKeyOffice = Registry.LocalMachine.OpenSubKey(string.Format(_addinOfficeRegistryKey, item.ToString()) + progID.Value, true);
                    else
                        regKeyOffice = Registry.CurrentUser.OpenSubKey(string.Format(_addinOfficeRegistryKey, item.ToString()) + progID.Value, true);

                    if (null == regKeyOffice)
                        continue;
                    regKeyOffice.SetValue(name, value);
                    regKeyOffice.Close();
                    //regKeyOffice.Dispose(); not available in previous .net versions
                }

                return true;
            }
            catch (Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (throwException)
                    throw;
                else
                    return false;
            }
        }

        #endregion

        #region Virtual Methods

        /// <summary>
        /// Create the used utils. The method was called in OnConnection
        /// </summary>
        /// <returns>new ToolsUtils instance</returns>
        protected internal virtual CommonUtils OnCreateUtils()
        {
            return new CommonUtils(this, 3 == _automationCode ? true : false, this.Type.Assembly);
        }

        /// <summary>
        /// Create the used factory. The method was called as first in the base ctor
        /// </summary>
        /// <returns>new Settings instance</returns>
        protected virtual Core CreateFactory()
        {
            return new Core();
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

        #region ErrorHandler
        
        /// <summary>
        /// Checks for a static method, signed with the ErrorHandlerAttribute and call them if its available
        /// </summary>
        /// <param name="type">type information for the class wtih static method </param>
       /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        private static void RaiseStaticErrorHandlerMethod(Type type, RegisterErrorMethodKind methodKind, NetRuntimeSystem.Exception exception)
        {
			MethodInfo errorMethod = AttributeHelper.GetRegisterErrorMethod(type);
            if (null != errorMethod)
                errorMethod.Invoke(null, new object[] { methodKind, exception });
        }

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
        [ComRegisterFunctionAttribute, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
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
				MultiRegisterAttribute attribute = MultiRegisterAttribute.GetAttribute(type);

                Assembly thisAssembly = Assembly.GetAssembly(type);
				string assemblyVersion = thisAssembly.GetName().Version.ToString();
                RegistryKey key = Registry.ClassesRoot.CreateSubKey("CLSID\\{" + type.GUID.ToString().ToUpper() + "}\\InprocServer32\\" + assemblyVersion);
                key.SetValue("CodeBase", thisAssembly.CodeBase);
                key.Close();

				Registry.ClassesRoot.CreateSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable");
				key = Registry.ClassesRoot.OpenSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\InprocServer32", true);
				key.SetValue("", NetRuntimeSystem.Environment.SystemDirectory + @"\mscoree.dll", RegistryValueKind.String);
				key.Close();

                // add bypass key
                // http://support.microsoft.com/kb/948461
                key = Registry.ClassesRoot.CreateSubKey("Interface\\{000C0601-0000-0000-C000-000000000046}");
                string defaultValue = key.GetValue("") as string;
                if (null == defaultValue)
                    key.SetValue("", "Office .NET Framework Lockback Bypass Key");
                key.Close();

				foreach(RegisterIn item in attribute.Products)
				{
					 // register addin in Office
                    if (location.Value == RegistrySaveLocation.LocalMachine)
                        Registry.LocalMachine.CreateSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value);
                    else
                        Registry.CurrentUser.CreateSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value);

					RegistryKey regKeyProduct = null;
					if(location.Value == RegistrySaveLocation.LocalMachine)
                        regKeyProduct = Registry.LocalMachine.OpenSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value, true);
					else
                        regKeyProduct = Registry.CurrentUser.OpenSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value, true);

					regKeyProduct.SetValue("LoadBehavior", addin.LoadBehavior);
					regKeyProduct.SetValue("FriendlyName", addin.Name);
					regKeyProduct.SetValue("Description", addin.Description);
					if(-1 != addin.CommandLineSafe)
						regKeyProduct.SetValue("CommandLineSafe", addin.CommandLineSafe);

					regKeyProduct.Close();
				}

                 if( (registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallAfter))
                        registerMethod.Invoke(null, new object[] { type, RegisterCall.CallAfter });
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception);
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
				MultiRegisterAttribute attribute = MultiRegisterAttribute.GetAttribute(type);

                // unregister addin
                Registry.ClassesRoot.DeleteSubKey(@"CLSID\{" + type.GUID.ToString().ToUpper() + @"}\Programmable", false);
               
			    foreach(RegisterIn item in attribute.Products)
				{
				    // unregister addin in office 
					if (location.Value == RegistrySaveLocation.LocalMachine)
                        Registry.LocalMachine.DeleteSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value, false);
					else
                        Registry.CurrentUser.DeleteSubKey(string.Format(_addinOfficeRegistryKey, MultiRegisterAttribute.RegistryEntry(item)) + progId.Value, false);
				}

                if ((registerMethodPresent) && (registerAttribute.Value == RegisterMode.CallBeforeAndAfter || registerAttribute.Value == RegisterMode.CallAfter))
                    registerMethod.Invoke(null, new object[] { type, RegisterCall.CallAfter });
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.UnRegister, exception);
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
        /// Try to detect the registry end for the current loaded host application(unkown)
        /// </summary>
        /// <param name="applicationProxy">application proy</param>
        /// <returns>Application name or null if failed</returns>
        private string TryDetectHostRegistryKey(object applicationProxy)
        {
            Guid libGuid = Factory.GetParentLibraryGuid(applicationProxy);
            if (Guid.Empty == libGuid)
                return null;
            string stringGuid = libGuid.ToString().ToUpper();
			stringGuid = stringGuid.Replace("{","").Replace("}","");

            switch (stringGuid)
            {
                case "00020813-0000-0000-C000-000000000046":
                    return "Excel";
                case "00020905-0000-0000-C000-000000000046":
                    return "Word";
                case "00062FFF-0000-0000-C000-000000000046":
                    return "Outlook";
                case "91493440-5A91-11CF-8700-00AA0060263B":
                    return "PowerPoint";
                case "4AFFC9A0-5F99-101B-AF4E-00AA003F0F07":
                    return "Access";
                case "A7107640-94DF-1068-855E-00DD01075445":
                    return "MSProject";
                default:
                    return null;
            }
        }

        #endregion
    }
}