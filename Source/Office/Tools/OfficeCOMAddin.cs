using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.IO;
using NetOffice.Tools;
using Office = NetOffice.OfficeApi;
using System.Runtime.CompilerServices;
using System.Linq;

namespace NetOffice.OfficeApi.Tools
{
    /// <summary>
    /// Represents common addin services
    /// </summary>
    /// <remarks>Some applications does not support the given features here in general or depending to its version</remarks>
    public abstract class OfficeCOMAddin : COMAddinBase, IOfficeCOMAddin
    {
        #region Fields

        /// <summary>
        /// OnStartupCompleteEvent Handler
        /// </summary>
        protected OnStartupCompleteEventHandler _onStartupCompleteEvent;

        /// <summary>
        /// OnDisconnectionEvent Handler
        /// </summary>
        protected OnDisconnectionEventHandler _onDisconnectionEvent;

        /// <summary>
        /// OnConnectionEvent Handler
        /// </summary>
        protected OnConnectionEventHandler _onConnectionEvent;

        /// <summary>
        /// OnAddInsUpdateEvent Handler
        /// </summary>
        protected OnAddInsUpdateEventHandler _onAddInsUpdateEvent;

        /// <summary>
        /// OnBeginShutdownEvent Handler
        /// </summary>
        protected OnBeginShutdownEventHandler _onBeginShutdownEvent;

        /// <summary>
        /// First field in OnConnection custom argument array
        /// </summary>
        protected int _automationCode = -1;

        /// <summary>
        /// Cache field used in IsLoadedFromSystem() method
        /// </summary>
        protected bool? _isLoadedFromSystem;

        #endregion

        #region Ctor

        /// <summary>
        /// Creates an instance of the class
        /// </summary>
        public OfficeCOMAddin() : base()
        {
            TaskPanes = new CustomTaskPaneCollection();
        }

        #endregion

        #region Properties

        /// <summary>
        /// Host Application Instance
        /// </summary>
        public virtual ICOMObject Application { get; protected set; }

        /// <summary>
        /// Ribbon instance to manipulate ui at runtime
        /// </summary>
        protected Office.IRibbonUI RibbonUI { get; set; }

        /// <summary>
        /// TaskPaneFactory to create custom task panes
        /// </summary>
        public Office.ICTPFactory TaskPaneFactory { get; set; }

        /// <summary>
        /// Collection with all created custom Task Panes
        /// </summary>
        public CustomTaskPaneCollection TaskPanes { get; private set; }

        /// <summary>
        /// ITaskPane Instances
        /// </summary>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Advanced)]
        protected IEnumerable<ITaskPane> TaskPaneInstances
        {
            get
            {
                List<ITaskPane> result = new List<ITaskPane>();
                foreach (var item in TaskPanes)
                {
                    ITaskPane match = item.Pane as ITaskPane;
                    if (null != match)
                        result.Add(match);
                }
                return result.ToArray();
            }
        }

        /// <summary>
        /// Custom addin object if created
        /// </summary>
        protected internal object CustomObject { get; private set; }

        /// <summary>
        /// Returns an enumerable sequence with instance managed com objects on root level
        /// </summary>
        /// <returns>ICOMObject enumerator</returns>
        [Browsable(false), EditorBrowsable(EditorBrowsableState.Never)]
        protected internal virtual IEnumerable<ICOMObject> OnCreateRoots()
        {
            List<ICOMObject> result = new List<ICOMObject>();
            if (null != Application)
                result.Add(Application);
            if (null != RibbonUI)
                result.Add(RibbonUI);
            if (null != TaskPaneFactory)
                result.Add(TaskPaneFactory);

            return result.ToArray();
        }

        #endregion

        #region IDTExtensibility2

        /// <summary>
        /// Occurs whenever an add-in is loaded into MS-Office
        /// </summary>
        /// <param name="application">A reference to an instance of the office application</param>
        /// <param name="connectMode">An ext_ConnectMode enumeration value that indicates the way the add-in was loaded into MS-Office</param>
        /// <param name="addInInst">An AddIn reference to the add-in's own instance. This is stored for later use, such as determining the parent collection for the add-in</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected internal abstract void HandleOnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is unloaded from MS Office
        /// </summary>
        /// <param name="removeMode">An ext_DisconnectMode enumeration value that informs an add-in why it was unloaded.</param>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use after the add-in unloads</param>
        protected internal abstract void HandleOnDisconnection(ext_DisconnectMode removeMode, ref Array custom);

        /// <summary>
        ///  Occurs whenever an add-in, which is set to load when MS Office starts, loads.
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use when the add-in loads</param>
        protected internal abstract void HandleOnStartupComplete(ref Array custom);

        /// <summary>
        /// Occurs whenever an add-in is loaded or unloaded from MS Office
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected internal abstract void HandleOnAddInsUpdate(ref Array custom);

        /// <summary>
        /// Occurs whenever MS Office shuts down while an add-in is running
        /// </summary>
        /// <param name="custom">An empty array that you can use to pass host-specific data for use in the add-in</param>
        protected internal abstract void HandleOnBeginShutdown(ref Array custom);

        void NetOffice.Tools.Native.IDTExtensibility2.OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                HandleOnConnection(application, connectMode, addInInst, ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnConnection, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                HandleOnDisconnection(removeMode, ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnStartupComplete(ref Array custom)
        {
            try
            {
                HandleOnStartupComplete(ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnAddInsUpdate(ref Array custom)
        {
            try
            {
                HandleOnAddInsUpdate(ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnAddInsUpdate, exception);
            }
        }

        void NetOffice.Tools.Native.IDTExtensibility2.OnBeginShutdown(ref Array custom)
        {
            try
            {
                HandleOnBeginShutdown(ref custom);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
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
        public event OnStartupCompleteEventHandler OnStartupComplete
        {
            add
            {
                _onStartupCompleteEvent += value;
            }
            remove
            {
                _onStartupCompleteEvent -= value;
            }
        }

        /// <summary>
        /// The Shutdown event occurs when the COM add-in is unloaded.
        /// You can use the OnDisconnection event procedure to run code that restores any changes made to the application by the add-in and to perform general clean-up operations.
        /// An add-in can be unloaded in one of the following ways:
        /// - The user clears the check box next to the add-in in the COM Add-ins dialog box.
        /// - The host application closes. If the add-in is loaded when the application closes, it is unloaded.
        ///   If the add-in's load behavior is set to Startup, it is reloaded when the application starts again.
        /// - The Connect property of the corresponding COMAddIn object is set to False.
        /// </summary>
        public event OnDisconnectionEventHandler OnDisconnection
        {
            add
            {
                _onDisconnectionEvent += value;
            }
            remove
            {
                _onDisconnectionEvent -= value;
            }
        }

        /// <summary>
        /// The OnConnection event occurs when the COM add-in is loaded (connected). An add-in can be loaded in one of the following ways:
        /// The user starts the host application and the add-in's load behavior is specified to load when the application starts.
        /// The user loads the add-in in the COM Add-ins dialog box.
        /// The Connect property of the corresponding COMAddIn object is set to True.
        /// For more information about the COMAddIn object, search the Microsoft® Office Visual Basic Reference Help index for "COMAddIn object."
        /// </summary>
        public event OnConnectionEventHandler OnConnection
        {
            add
            {
                _onConnectionEvent += value;
            }
            remove
            {
                _onConnectionEvent -= value;
            }
        }

        /// <summary>
        /// The OnAddInsUpdate event occurs when the set of loaded COM add-ins changes.
        /// When an add-in is loaded or unloaded, the OnAddInsUpdate event occurs in any other loaded add-ins.
        /// For example, if add-ins A and B both are loaded currently, and then add-in C is loaded,
        /// the OnAddInsUpdate event occurs in add-ins A and B. If C is unloaded, the OnAddInsUpdate event occurs again in add-ins A and B.
        /// </summary>
        public event OnAddInsUpdateEventHandler OnAddInsUpdate
        {
            add
            {
                _onAddInsUpdateEvent += value;
            }
            remove
            {
                _onAddInsUpdateEvent -= value;
            }
        }

        /// <summary>
        /// The OnBeginShutdown event occurs when the host application begins its shutdown routines,
        /// in the case where the application closes while the COM add-in is still loaded.
        /// If the add-in is not loaded when the application closes,
        /// the OnBeginShutdown event does not occur. When this event does occur, it occurs before the OnDisconnection event.
        /// You can use the OnBeginShutdown event procedure to run code when the user closes the application. For example, you can run code that saves form data to a file.
        /// </summary>
        public event OnBeginShutdownEventHandler OnBeginShutdown
        {
            add
            {
                _onBeginShutdownEvent += value;
            }
            remove
            {
                _onBeginShutdownEvent -= value;
            }
        }

        #endregion

        #region IRibbonExtensibility

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="RibbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        string Office.Native.IRibbonExtensibility.GetCustomUI(string RibbonID)
        {
            try
            {
                return OnGetCustomUI(RibbonID);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
                return string.Empty;
            }
        }

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="ribbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        protected internal virtual string OnGetCustomUI(string ribbonID)
        {
            var ribbon = NetOffice.Attributes.AttributeExtensions.GetCustomAttribute<CustomUIAttribute>(Type);
            if (null != ribbon && ribbon.RibbonIDs.Contains(ribbonID))
                return ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
            else
                return string.Empty;
        }

        /// <summary>
        /// Pre-defined Ribbon Loader
        /// </summary>
        /// <param name="ribbonUI">actual ribbon ui</param>
        public virtual void CustomUI_OnLoad(Office.Native.IRibbonUI ribbonUI)
        {
            try
            {
                RibbonUI = Factory.CreateKnownObjectFromComProxy<OfficeApi.IRibbonUI>(null, ribbonUI, typeof(OfficeApi.IRibbonUI));
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
            }
        }

        #endregion

        #region ICustomTaskPaneConsumer

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="CTPFactoryInst">factory proxy from host application</param>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        void Office.Native.ICustomTaskPaneConsumer.CTPFactoryAvailable(object CTPFactoryInst)
        {
            try
            {
                if (null == CTPFactoryInst)
                {
                    Factory.Console.WriteLine("Warning: null argument recieved in CTPFactoryAvailable. argument name: CTPFactoryInst");
                    return;
                }

                OfficeApi.ICTPFactory taskPaneFactory = Factory.CreateKnownObjectFromComProxy<NetOffice.OfficeApi.ICTPFactory>(null, CTPFactoryInst, typeof(NetOffice.OfficeApi.ICTPFactory));
                OnCTPFactoryAvailable(taskPaneFactory);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.CTPFactoryAvailable, exception);
            }
        }

        /// <summary>
        /// ICustomTaskPaneConsumer implementation
        /// </summary>
        /// <param name="ctpFactoryInst">factory proxy from host application</param>
        protected internal virtual void OnCTPFactoryAvailable(OfficeApi.ICTPFactory ctpFactoryInst)
        {
            CustomTaskPaneHandler paneHandler = new CustomTaskPaneHandler();
            paneHandler.ProceedCustomPaneAttributes(Factory, ctpFactoryInst, TaskPanes, OnError, Application, Type, this, CallOnCreateTaskPaneInfo, AttributePane_VisibleStateChange, AttributePane_DockPositionStateChange);
            paneHandler.CreateCustomPanes(Factory, ctpFactoryInst, TaskPanes, OnError, Application);
        }

        /// <summary>
        /// The method is called while the CustomPane attribute is processed
        /// </summary>
        /// <param name="paneInfo">pane definition</param>
		/// <returns>true if pane should be create, otherwise false</returns>
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
                foreach (TaskPaneInfo item in TaskPanes)
                {
                    if (item.Pane.UnderlyingObject == customTaskPaneInst.UnderlyingObject)
                    {
                        try
                        {
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
                            if (null != target)
                            {
                                try
                                {
                                    target.OnVisibleStateChanged(item.Pane.Visible);
                                }
                                catch (Exception exception)
                                {
                                    Factory.Console.WriteException(exception);
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            Factory.Console.WriteException(exception);
                        }
                    }
                }
                TaskPaneVisibleStateChanged(customTaskPaneInst);
            }
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
            }
        }

        private void CallTaskPaneDockPositionStateChange(NetOffice.OfficeApi._CustomTaskPane customTaskPaneInst)
        {
            try
            {
                foreach (TaskPaneInfo item in TaskPanes)
                {
                    if (item.Pane.UnderlyingObject == customTaskPaneInst.UnderlyingObject)
                    {
                        try
                        {
                            ITaskPane target = item.Pane.ContentControl as ITaskPane;
                            if (null != target)
                            {
                                try
                                {
                                    target.OnDockPositionChanged(item.Pane.DockPosition);
                                }
                                catch (Exception exception)
                                {
                                    Factory.Console.WriteException(exception);
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            Factory.Console.WriteException(exception);
                        }
                    }
                }
                TaskPaneDockStateChanged(customTaskPaneInst);
            }
            catch (Exception exception)
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
            catch (Exception exception)
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
            catch (Exception exception)
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
            catch (Exception exception)
            {
                Factory.Console.WriteException(exception);
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

        #region Methods

        /// <summary>
        /// Raise the OnStartupComplete event
        /// </summary>
        /// <param name="custom">custom arguments</param>
        protected internal virtual void RaiseOnStartupComplete(ref Array custom)
        {
            try
            {
                var handler = _onStartupCompleteEvent;
                if (null != handler)
                    handler(ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnStartupComplete, exception);
            }
        }

        /// <summary>
        /// Raise the OnDisconnection event
        /// </summary>
        /// <param name="removeMode">kind of remove</param>
        /// <param name="custom">custom arguments</param>
        protected internal virtual void RaiseOnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                var handler = _onDisconnectionEvent;
                if (null != handler)
                    handler(removeMode, ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnDisconnection, exception);
            }
        }

        /// <summary>
        /// Raise the OnConnection event
        /// </summary>
        /// <param name="application">application host instance</param>
        /// <param name="connectMode">kind of connect</param>
        /// <param name="addInInst">addin instance</param>
        /// <param name="custom">custom arguments</param>
        protected internal virtual void RaiseOnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            try
            {
                var handler = _onConnectionEvent;
                if (null != handler)
                    handler(Application, connectMode, addInInst, ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnConnection, exception);
            }
        }

        /// <summary>
        /// Raise the OnAddInsUpdate event
        /// </summary>
        /// <param name="custom">custom arguments</param>
        protected internal virtual void RaiseOnAddInsUpdate(ref Array custom)
        {
            try
            {
                var handler = _onAddInsUpdateEvent;
                if (null != handler)
                    handler(ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnAddInsUpdate, exception);
            }
        }

        /// <summary>
        /// Raise the OnBeginShutdown event
        /// </summary>
        /// <param name="custom">custom arguments</param>
        protected internal virtual void RaiseOnBeginShutdown(ref Array custom)
        {
            try
            {
                var handler = _onBeginShutdownEvent;
                if (null != handler)
                    handler(ref custom);
            }
            catch (System.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.OnBeginShutdown, exception);
            }
        }

        /// <summary>
        /// Try to detect the addin is loaded from system hive key
        /// The method does not work for non specific application addins
        /// </summary>
        /// <returns>null if unkown or true/false</returns>
        protected internal bool? IsLoadedFromSystem()
        {
            if (null != _isLoadedFromSystem)
                return _isLoadedFromSystem;

            OfficeApi.Tools.Contribution.RegistryLocationResult result =
                OfficeApi.Tools.Contribution.CommonUtils.TryFindAddinLoadLocation(Type,
                                        ApplicationIdentifiers.ApplicationType.PowerPoint);
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
        protected internal void TryCreateCustomObject(object addInInst)
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
        /// Read string from resource in executing assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <returns>resource string</returns>
        private string ReadString(string resourceAddress)
        {
            return ReadString(resourceAddress, Type.Assembly);
        }

        /// <summary>
        /// Read string from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource string</returns>
        private string ReadString(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = ReadStream(resourceAddress, assembly);
            System.IO.StreamReader textStreamReader = new System.IO.StreamReader(resourceStream);
            if (textStreamReader == null)
                throw (new System.IO.IOException("Error accessing resource string."));

            string text = textStreamReader.ReadToEnd();
            textStreamReader.Close();
            resourceStream.Close();
            return text;
        }

        /// <summary>
        /// Read stream from resource in given assembly
        /// </summary>
        /// <param name="resourceAddress">full qualified resource address</param>
        /// <param name="assembly">target assembly to find resource</param>
        /// <returns>resource stream</returns>
        public Stream ReadStream(string resourceAddress, Assembly assembly)
        {
            System.IO.Stream resourceStream = assembly.GetManifestResourceStream(resourceAddress);
            if (resourceStream == null)
            {
                string target = Type.Namespace + "." + resourceAddress;
                resourceStream = assembly.GetManifestResourceStream(target);
            }

            if (resourceStream == null)
                throw (new System.IO.IOException("Error accessing resource Stream."));

            return resourceStream;
        }

        #endregion
    }
}
