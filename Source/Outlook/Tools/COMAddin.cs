using System;
using NetRuntimeSystem = System;
using System.Collections;
using System.Collections.Generic;
using System.Reflection;
using Microsoft.Win32;
using System.ComponentModel;
using System.Runtime.InteropServices;
using NetOffice;
using NetOffice.Attributes;
using NetOffice.Tools;
using NetOffice.OfficeApi.Tools;
using Office = NetOffice.OfficeApi;
using Outlook = NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using System.Runtime.CompilerServices;

namespace NetOffice.OutlookApi.Tools
{  
    /// <summary>
    /// NetOffice MS-Outlook COM Addin
    /// </summary>
	[ComVisible(true), ClassInterface(ClassInterfaceType.AutoDual)]
    public abstract class COMAddin : COMAddinBase, IOfficeCOMAddin, Native.FormRegionStartup
    {
        #region Fields

        /// <summary>
        /// MS-Outlook Addin Registry Path 
        /// </summary>
        private static readonly string _addinOfficeRegistryKey  = "Software\\Microsoft\\Office\\Outlook\\Addins\\";

        /// <summary>
        /// MS-Outlook FormRegion Registry Path 
        /// </summary>
        private static readonly string _formRegionsOfficeRegistryKey = "Software\\Microsoft\\Office\\Outlook\\FormRegions\\";

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
            OpenFormRegions = new List<OpenFormRegion>();
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
        protected internal Outlook.Application Application { get; private set; }
        
        /// <summary>
        /// Collection with all created custom Task Panes
        /// </summary>
        protected CustomTaskPaneCollection TaskPanes { get; private set; }

        /// <summary>
        /// TaskPaneFactory from CTPFactoryAvailable
        /// </summary>
        public Office.ICTPFactory TaskPaneFactory { get; set; }

		/// <summary>
        /// ITaskPane Instances
        /// </summary>
		protected List<ITaskPane> TaskPaneInstances { get; set; }

        /// <summary>
        /// Ribbon instance to manipulate ui at runtime 
        /// </summary>
        protected Office.IRibbonUI RibbonUI { get; private set; }

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
        /// Shutdown Changes for Outlook 2010: https://msdn.microsoft.com/library/office/ee720183.aspx
        /// ------------------------------------------------------------------------------------------
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

        #region IDTExtensibility2 Members

        void NetOffice.Tools.Native.IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                Tweaks.ApplyTweaks(Factory, this, Type, "Outlook", IsLoadedFromSystem);
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

        void NetOffice.Tools.Native.IDTExtensibility2.OnConnection(object application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try
            {
                if (null != custom && custom.Length > 0)
                {
                    object firstCustomItem = custom.GetValue(1);
                    string tryString = null != firstCustomItem ? firstCustomItem.ToString() : String.Empty;
                    NetRuntimeSystem.Int32.TryParse(tryString, out _automationCode);
                }

                this.Application = new Outlook.Application(Factory, null, application);
                Utils = OnCreateUtils();
                TryCreateCustomObject(AddInInst);
                RaiseOnConnection(this.Application, ConnectMode, AddInInst, ref custom);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
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
                        NetOffice.DebugConsole.Default.WriteException(exception);
                    }
                }

                try
                {
                    foreach (var item in OpenFormRegions)
                    {
                        try
                        {
                            IDisposable disposable = item as IDisposable;
                            if (null != disposable)
                                disposable.Dispose();
                            else
                                item.UnderlyingRegion.Dispose();
                        }
                        catch (NetRuntimeSystem.Exception exception)
                        {
                            NetOffice.DebugConsole.Default.WriteException(exception);
                        }
                    }
                    OpenFormRegions.Clear();
                }
                catch (NetRuntimeSystem.Exception exception)
                {
                    NetOffice.DebugConsole.Default.WriteException(exception);
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
                    if (null != RibbonUI)
                    {
                        RibbonUI.Dispose();
                        RibbonUI = null;
                    }
                }
                catch (NetRuntimeSystem.Exception exception)
                {
                    NetOffice.DebugConsole.Default.WriteException(exception);
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

        void NetOffice.Tools.Native.IDTExtensibility2.OnAddInsUpdate(ref Array custom)
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

        void NetOffice.Tools.Native.IDTExtensibility2.OnBeginShutdown(ref Array custom)
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
        /// <param name="RibbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual string GetCustomUI(string RibbonID)
        {
            try
            {
                OlCustomUIAttribute olRibbon = GetOlRibbonAttribute(Type, RibbonID);
                if (null != olRibbon)
                {
                    return Utils.Resource.ReadString(OlCustomUIAttribute.BuildPath(olRibbon.Value, olRibbon.UseAssemblyNamespace, Type.Namespace));
                }
                else
                {
                    CustomUIAttribute ribbon = AttributeReflector.GetRibbonAttribute(Type, RibbonID);
                    if (null != ribbon)
                        return Utils.Resource.ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
                    else
                        return String.Empty;
                }
            }
            catch (NetRuntimeSystem.Exception exception)
            {
				NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
                return String.Empty;
            } 
        }

        /// <summary>
        /// Pre-defined Ribbon Loader
        /// </summary>
        /// <param name="ribbonUI">actual ribbon ui</param>
        public virtual void CustomUI_OnLoad(Office.Native.IRibbonUI ribbonUI)
        {
            try
            {
                RibbonUI = COMObject.Create<OfficeApi.IRibbonUI>(Factory, ribbonUI);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
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
                paneHandler.ProceedCustomPaneAttributes(TaskPanes, Type, this, CallOnCreateTaskPaneInfo, AttributePane_VisibleStateChange, AttributePane_DockPositionStateChange);
                TaskPaneFactory = paneHandler.CreateCustomPanes<ITaskPane, Outlook.Application>(Factory, CTPFactoryInst, TaskPanes, TaskPaneInstances, OnError, Application);
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
		/// <returns>true if pane shoul be create, otherwise false</returns>
		protected internal virtual bool OnCreateTaskPaneInfo(TaskPaneInfo paneInfo)
		{
			return true;
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
								catch(NetRuntimeSystem.Exception exception)
								{
									Factory.Console.WriteException(exception);
								}
							}
						}
						catch(NetRuntimeSystem.Exception exception)
						{
							Factory.Console.WriteException(exception);
						}
					}
				}
                TaskPaneVisibleStateChanged(customTaskPaneInst);
			}
			catch(NetRuntimeSystem.Exception exception)
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
								catch(NetRuntimeSystem.Exception exception)
								{
									Factory.Console.WriteException(exception);
								}
							}
						}
						catch(NetRuntimeSystem.Exception exception)
						{
							Factory.Console.WriteException(exception);
						}
					}
				}
                TaskPaneDockStateChanged(customTaskPaneInst);
			}
			catch(NetRuntimeSystem.Exception exception)
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
			catch(NetRuntimeSystem.Exception exception)
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
			catch(NetRuntimeSystem.Exception exception)
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
			catch(NetRuntimeSystem.Exception exception)
			{
				Factory.Console.WriteException(exception);
			}            
        }

        #endregion

        #region FormRegionStartup

        /// <summary>
        /// Current open Form Regions
        /// </summary>
        protected List<OpenFormRegion> OpenFormRegions { get; private set; }

        /// <summary>
        /// Occurs after a form region has been opened
        /// </summary>
        public event FormRegionEventHandler FormRegionOpen;

        /// <summary>
        /// Occurs after a form region has been closed
        /// </summary>
        public event FormRegionEventHandler FormRegionClose;

        /// <summary>
        /// Raise the FormRegionOpen event
        /// </summary>
        /// <param name="form"></param>
        protected virtual void OnFormRegionOpen(OpenFormRegion form)
        {
            FormRegionOpen?.Invoke(form);
        }

        /// <summary>
        /// Raise the FormRegionClose event
        /// </summary>
        /// <param name="form"></param>
        protected virtual void OnFormRegionClose(OpenFormRegion form)
        {
            FormRegionClose?.Invoke(form);
        }

        /// <summary>
        /// Creates an new instance of OpenFormRegion
        /// </summary>
        /// <param name="form">underlying form region</param>
        /// <returns>new instance of OpenFormRegion</returns>
        protected virtual OpenFormRegion OnCreateOpenFormRegion(FormRegion form)
        {
            OpenFormRegion openForm = new OpenFormRegion(form);
            return openForm;
        }

        /// <summary>
        /// Obtains appropriate storage for a form region based on the specified information.
        /// </summary>
        /// <param name="formRegionName">The internal name of the form region. This can be indicated by the name tag in the corresponding form region XML manifest.</param>
        /// <param name="item">The Outlook item object that caused the loading of the form region.</param>
        /// <param name="lcid">The current locale ID.</param>
        /// <param name="formRegionMode">The mode that the form region is being loaded into.</param>
        /// <param name="formRegionSize">The type of form region being loaded, either adjoining or separate.</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual object GetFormRegionStorage(object formRegionName, object item, object lcid, object formRegionMode, object formRegionSize)
        {
            try
            {
                FormRegionAttribute attribute = FormRegionAttribute.GetAttribute(Type, (string)formRegionName, (int)lcid);
                if (null != attribute)
                    return Utils.Resource.ReadBytes(CustomUIAttribute.BuildPath(attribute.StorageFile, true, Type.Namespace));
                else
                    return null;
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                OnOutlookError(OutlookErrorMethodKind.GetFormRegionStorage, exception);
                return null;
            }
        }

        /// <summary>
        /// Allows an add-in to update the user interface of a form region before it is displayed. 
        /// </summary>
        /// <param name="formRegion">The FormRegion object representing the form region that is to be displayed</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual void BeforeFormRegionShow(object formRegion)
        {
            try
            {
                FormRegion form = new Outlook.FormRegion(null, formRegion);
                OpenFormRegion openForm = OnCreateOpenFormRegion(form);
                if (null == openForm)
                    openForm = new OpenFormRegion(form);
                openForm.Close += OpenForm_Close;
                OpenFormRegions.Add(openForm);
                OnFormRegionOpen(openForm);
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                OnOutlookError(OutlookErrorMethodKind.BeforeFormRegionShow, exception);
            }
        }

        /// <summary>
        /// Obtains the XML manifest for a form region.
        /// </summary>
        /// <param name="FormRegionName">The name of the form region which is the name used when registering the form region in the Windows registry. </param>
        /// <param name="LCID">The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual object GetFormRegionManifest([MarshalAs(19)] [In] string FormRegionName, [In] int LCID)
        {
            try
            {
                FormRegionAttribute attribute = FormRegionAttribute.GetAttribute(Type, (string)FormRegionName, (int)LCID);
                if (null != attribute)
                    return Utils.Resource.ReadString(CustomUIAttribute.BuildPath(attribute.ManifestFile, true, Type.Namespace));
                else
                    return null;
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                OnOutlookError(OutlookErrorMethodKind.GetFormRegionManifest, exception);
                return null;
            }
        }

        /// <summary>
        /// Obtains an icon image that will be displayed for a particular type of icon for the form region.
        /// </summary>
        /// <param name="formRegionName">The name of the form region which is the name used when registering the form region in the Windows registry. </param>
        /// <param name="lcid">The locale ID that identifies the language that Outlook is currently using. This value is used to obtain the localization strings corresponding to this language for the form region.</param>
        /// <param name="icon">A constant that identifies the type of icon.</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public virtual object GetFormRegionIcon(object formRegionName, object lcid, object icon)
        {
            try
            {
                FormRegionAttribute attribute = FormRegionAttribute.GetAttribute(Type, (string)formRegionName, (int)lcid);
                if (null != attribute)
                {
                    Enums.OlFormRegionIcon olIcon = (Enums.OlFormRegionIcon)icon;
                    if (attribute.OlIconWildcard)
                    {
                        var readIcon = Utils.Resource.ReadIcon(attribute.IconFile);
                        return Utils.Image.ToPicture(readIcon);
                    }
                    if (attribute.OlIcon == olIcon)
                    {
                        var readIcon = Utils.Resource.ReadIcon(attribute.IconFile);
                        return Utils.Image.ToPicture(readIcon);
                    }
                    return null;
                }
                else
                    return null;
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                OnOutlookError(OutlookErrorMethodKind.GetFormRegionIcon, exception);
                return null;
            }
        }

        private void OpenForm_Close(OpenFormRegion form)
        {
            try
            {
                OpenFormRegions.Remove(form);
                OnFormRegionClose(form);
                IDisposable disposable = form as IDisposable;
                if (null != disposable)
                    disposable.Dispose();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                OnOutlookError(OutlookErrorMethodKind.CloseOpenFormRegion, exception);
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
                                                          ApplicationIdentifiers.ApplicationType.Outlook,
                                                          addinType,
                                                          name, value,
                                                          throwException);
        }

        #endregion

        #region Virtual Methods

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
        /// <returns>new Settings instance</returns>
        protected virtual Core CreateFactory()
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
        /// Looks for the CustomUIAttribute
        /// </summary>
        /// <param name="type">the type you want looking for the attribute</param>
        /// <param name="ribbonID">target window id</param>
        /// <returns>CustomUIAttribute or null</returns>
        private static OlCustomUIAttribute GetOlRibbonAttribute(Type type, string ribbonID)
        {
            object[] array = type.GetCustomAttributes(typeof(OlCustomUIAttribute), false);
            if (array.Length == 0)
                return null;

            foreach (OlCustomUIAttribute item in array)
            {
                if (item.RibbonID.Equals(ribbonID, StringComparison.InvariantCultureIgnoreCase))
                    return item;
            }

            return null;
        }

        /// <summary>
        /// Try to detect the addin is loaded from system hive key
        /// </summary>
        /// <returns>null if unkown or true/false</returns>
        private bool? IsLoadedFromSystem()
        {
            if (null != _isLoadedFromSystem)
                return _isLoadedFromSystem;

            OfficeApi.Tools.Contribution.RegistryLocationResult result = 
                OfficeApi.Tools.Contribution.CommonUtils.TryFindAddinLoadLocation(Type,
                                        ApplicationIdentifiers.ApplicationType.Outlook);
            switch (result)
            {
                case Office.Tools.Contribution.RegistryLocationResult.User:
                    _isLoadedFromSystem = false;
                    break;
                case Office.Tools.Contribution.RegistryLocationResult.System:
                    _isLoadedFromSystem = true;
                    break;
                //default:
                //    throw new IndexOutOfRangeException();
            }

            return _isLoadedFromSystem;
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
                    addInInst.GetType().InvokeMember("Object", NetRuntimeSystem.Reflection.BindingFlags.SetProperty, null, addInInst, param);
                }
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CreateCustomAddinInstance, exception);
            }
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

        /// <summary>
        /// Custom outlook-specific error handler
        /// </summary>
        /// <param name="methodKind">origin method where the error comes from</param>
        /// <param name="exception">occured exception</param>
        protected virtual void OnOutlookError(OutlookErrorMethodKind methodKind, NetRuntimeSystem.Exception exception)
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
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            COMAddinRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeRegisterKeyState.NeedToCreate);
            RegisterHandleRequireShutdownNotificationAttribute(type);
            RegisterHandleFormRegionAttribute(type);
        }

        /// <summary>
        /// Called from regasm while ungregister
        /// </summary>
        /// <param name="type">Type information for the class</param>
        [ComUnregisterFunctionAttribute, Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        public static void UnregisterFunction(Type type)
        {
            if (null == type)
                throw new ArgumentNullException("type");
            if (null != type.GetCustomAttribute<DontRegisterAddinAttribute>())
                return;

            COMAddinUnRegisterHandler.Proceed(type, new string[] { _addinOfficeRegistryKey }, InstallScope.System, OfficeUnRegisterKeyState.NeedToDelete);
            UnregisterHandleFormRegionAttribute(type);
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
            RegisterHandleRequireShutdownNotificationAttribute(type);
            RegisterHandleFormRegionAttribute(type);
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

            UnregisterHandleFormRegionAttribute(type);
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

        private static void RegisterHandleRequireShutdownNotificationAttribute(Type type)
        {
            try
            {
                if (null != RequireShutdownNotificationAttribute.GetAttribute(type))
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystem = location.IsMachineAddinTarget();
                    ProgIdAttribute progId = AttributeReflector.GetProgIDAttribute(type);
                    RequireShutdownNotificationAttribute.CreateApplicationKey(isSystem, _addinOfficeRegistryKey, progId.Value);
                }

            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if (!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        private static void RegisterHandleFormRegionAttribute(Type type)
        {
            try
            {
                var formAttributes = FormRegionAttribute.GetAttributes(type);
                foreach (var item in formAttributes)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystem = location.IsMachineAddinTarget();
                    ProgIdAttribute progId = AttributeReflector.GetProgIDAttribute(type);
                    FormRegionAttribute.CreateKey(isSystem, _formRegionsOfficeRegistryKey, progId.Value, item.Category, item.Name);
                }
            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if(!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        private static void UnregisterHandleFormRegionAttribute(Type type)
        {
            try
            {
                var formAttributes = FormRegionAttribute.GetAttributes(type);
                foreach (var item in formAttributes)
                {
                    RegistryLocationAttribute location = AttributeReflector.GetRegistryLocationAttribute(type);
                    bool isSystem = location.IsMachineAddinTarget();
                    ProgIdAttribute progId = AttributeReflector.GetProgIDAttribute(type);
                    FormRegionAttribute.TryDeleteKey(isSystem, _formRegionsOfficeRegistryKey, progId.Value, item.Category, item.Name);
                }

            }
            catch (System.Exception exception)
            {
                NetOffice.DebugConsole.Default.WriteException(exception);
                if(!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        #endregion
    }
}