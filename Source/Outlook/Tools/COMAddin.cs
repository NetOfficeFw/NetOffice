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
    public abstract class COMAddin : OfficeCOMAddin, IOfficeCOMAddin, Native.FormRegionStartup
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
        public new Outlook.Application Application
        {
            get
            {
                return base.Application as Outlook.Application;
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
                NetRuntimeSystem.Int32.TryParse(tryString, out _automationCode);
            }

            base.Application = Factory.CreateKnownObjectFromComProxy<Outlook.Application>(null, application, typeof(Outlook.Application));
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
                        Factory.Console.WriteException(exception);
                    }
                }
                OpenFormRegions.Clear();
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
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

        #region IRibbonExtensibility

        /// <summary>
        /// IRibbonExtensibility implementation
        /// </summary>
        /// <param name="ribbonID">target ribbon id</param>
        /// <returns>XML content or String.Empty</returns>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        protected override string OnGetCustomUI(string ribbonID)
        {
            try
            {
                OlCustomUIAttribute olRibbon = GetOlRibbonAttribute(Type, ribbonID);
                if (null != olRibbon)
                {
                    return Utils.Resource.ReadString(OlCustomUIAttribute.BuildPath(olRibbon.Value, olRibbon.UseAssemblyNamespace, Type.Namespace));
                }
                else
                {
                    var ribbon = NetOffice.Attributes.AttributeExtensions.GetCustomAttribute<CustomUIAttribute>(Type);
                    if (null != ribbon && CustomUIAttribute.ContainsProcessedRibbonId(ribbon, ribbon.RibbonID))
                        return Utils.Resource.ReadString(CustomUIAttribute.BuildPath(ribbon.Value, ribbon.UseAssemblyNamespace, Type.Namespace));
                    else
                        return string.Empty;
                }
            }
            catch (NetRuntimeSystem.Exception exception)
            {
                Factory.Console.WriteException(exception);
                OnError(ErrorMethodKind.GetCustomUI, exception);
                return String.Empty;
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
                FormRegion form = COMObject.Create<Outlook.FormRegion>(Factory, formRegion);
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

        #endregion

        #region ErrorHandler

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
                DebugConsole.Default.WriteException(exception);
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
                DebugConsole.Default.WriteException(exception);
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
                DebugConsole.Default.WriteException(exception);
                if(!RegisterErrorHandler.RaiseStaticErrorHandlerMethod(type, RegisterErrorMethodKind.Register, exception))
                    throw;
            }
        }

        #endregion
    }
}