using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.CoreServices;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.PowerPointApi.Behind
{
    /// <summary>
    /// CoClass Application
    /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745704.aspx </remarks>
    [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass), ComProgId("PowerPoint.Application"), ModuleProvider(typeof(ModulesLegacy.ApplicationModule))]
    [ComEventContract(typeof(NetOffice.PowerPointApi.EventContracts.EApplication))]
    [HasInteropCompatibilityClass(typeof(ApplicationClass))]
    public class Application : _Application, NetOffice.PowerPointApi.Application, IAutomaticQuit, IApplicationVersionProvider
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private NetOffice.PowerPointApi.Behind.EventContracts.EApplication_SinkHelper _eApplication_SinkHelper;

        private bool _versionRequested;
        private object _cachedVersion;
        private object _chachedVersionLock = new object();

        #endregion

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.PowerPointApi.Application);
                return _contractType;
            }
        }
        private static Type _contractType;


        /// <summary>
        /// Instance Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type InstanceType
        {
            get
            {
                return LateBindingApiWrapperType;
            }
        }

        /// <summary>
        /// Type Cache
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Application);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Application() : base()
        {

        }

        /// <summary>
        /// Creates a new instance of the class
        /// <param name="factory">factory core instead of default core</param>
        /// <param name="tryProxyServiceFirst">try to get a running application first before create a new application</param>
        /// </summary>
        public Application(Core factory = null, bool tryProxyServiceFirst = false) : base()
        {
            object proxy = null;
            if (tryProxyServiceFirst)
            {
                proxy = ProxyService.GetActiveInstance("PowerPoint", "Application", false);
                if (null != proxy)
                {
                    CreateFromProxy(proxy, true);
                    FromProxyService = true;
                }
            }

            if(null == proxy)
            {
                CreateFromProgId("PowerPoint.Application", true);
            }

            Factory = null != factory ? factory : Core.Default;
            TryRequestVersion();
            RegisterAsApplicationVersionProvider();
            OnCreate();
            _isInitialized = true;
        }

        #endregion

        #region ICloneable<Application>

        /// <summary>
        /// Creates a new Application that is a copy of the current instance
        /// </summary>
        /// <returns>A new Application that is a copy of this instance</returns>
        /// <exception cref="CloneException">An unexpected error occured. See inner exception(s) for details.</exception>
        public new virtual NetOffice.PowerPointApi.Application DeepCopy()
        {
            return base.Clone() as NetOffice.PowerPointApi.Application;
        }

        #endregion

        #region ICOMObjectProxyService

        /// <summary>
        /// Instance is created from an already running application
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced)]
        public bool FromProxyService { get; private set; }

        #endregion

        #region IApplicationVersionProvider

        string IApplicationVersionProvider.Name
        {
            get
            {
                return "Microsoft PowerPoint";
            }
        }

        string IApplicationVersionProvider.ComponentName
        {
            get
            {
                return "NetOffice.PowerPointApi";
            }
        }

        /// <summary>
        /// Request version information on demand and cache to call the remote server only 1x times
        /// </summary>
        object IApplicationVersionProvider.Version
        {
            get
            {
                lock (_chachedVersionLock)
                {
                    if (null == _cachedVersion)
                    {
                        _cachedVersion = TryVersionPropertyGet();
                    }
                }
                return _cachedVersion;
            }
        }

        bool IApplicationVersionProvider.VersionRequested
        {
            get
            {
                return _versionRequested;
            }
        }

        void IApplicationVersionProvider.TryRequestVersion()
        {
            _cachedVersion = TryVersionPropertyGet();
        }

        /// <summary>
        /// Try get version information without fail
        /// </summary>
        /// <returns></returns>
        private object TryVersionPropertyGet()
        {
            try
            {
                if (null != _proxyShare)
                    return Invoker.PropertyGet(this, "Version");
                else
                    return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                if (null != _proxyShare)
                    _versionRequested = true;
            }
        }

        #endregion

        #region IEventBinding

        /// <summary>
        /// Creates active sink helper
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void CreateEventBridge()
        {
            if (false == Factory.Settings.EnableEvents)
                return;

            if (null != _connectPoint)
                return;

            if (null == _activeSinkId)
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.PowerPointApi.Behind.EventContracts.EApplication_SinkHelper.Id);


            if (NetOffice.PowerPointApi.Behind.EventContracts.EApplication_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _eApplication_SinkHelper = new NetOffice.PowerPointApi.Behind.EventContracts.EApplication_SinkHelper(this, _connectPoint);
                return;
            }
        }

        /// <summary>
        /// The instance use currently an event listener
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get
            {
                return (null != _connectPoint);
            }
        }
        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <returns>true if one or more event is active, otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetCountOfEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Raise an instance event
        /// </summary>
        /// <param name="eventName">name of the event without 'Event' at the end</param>
        /// <param name="paramsArray">custom arguments for the event</param>
        /// <returns>count of called event recipients</returns>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
        {
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
        }
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
            if (null != _eApplication_SinkHelper)
            {
                _eApplication_SinkHelper.Dispose();
                _eApplication_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running PowerPoint.Application instances from the environment/system
        /// </summary>
        /// <returns>PowerPoint.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("PowerPoint", "Application");
        }

        /// <summary>
        /// Returns all running PowerPoint.Application instances from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>PowerPoint.Application sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances(Func<Application, bool> predicate)
        {
            return ProxyService.GetActiveInstances<Application>("PowerPoint", "Application", predicate);
        }

        /// <summary>
        /// Returns the count of running PowerPoint.Application instances that passed a predicate filter
        /// </summary>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount()
        {
            var sequence = ProxyService.GetActiveInstances<Application>("PowerPoint", "Application");
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns the count of running PowerPoint.Application instances that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <returns>count of running application</returns>
        public static int GetActiveInstancesCount(Func<Application, bool> predicate)
        {
            var sequence = ProxyService.GetActiveInstances<Application>("PowerPoint", "Application", predicate);
            int result = sequence.Count;
            sequence.Dispose();
            return result;
        }

        /// <summary>
        /// Returns a running PowerPoint.Application instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>PowerPoint.Application instance or null(Nothing in Visual Basic)</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("PowerPoint", "Application", throwExceptionIfNotFound);
        }

        /// <summary>
        /// Returns first running PowerPoint.Application instance from the environment/system that passed a predicate filter
        /// </summary>
        /// <param name="predicate">filter predicate</param>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>PowerPoint.Application instance or null(Nothing in Visual Basic)</returns>
        /// <exception cref="ArgumentOutOfRangeException">occurs if no instance match and throwExceptionIfNotFound is set</exception>
        public static Application GetActiveInstance(Func<Application, bool> predicate, bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("PowerPoint", "Application", predicate, throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_WindowSelectionChangeEventHandler _WindowSelectionChangeEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743918.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_WindowSelectionChangeEventHandler WindowSelectionChangeEvent
        {
            add
            {
                CreateEventBridge();
                _WindowSelectionChangeEvent += value;
            }
            remove
            {
                _WindowSelectionChangeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_WindowBeforeRightClickEventHandler _WindowBeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746559.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_WindowBeforeRightClickEventHandler WindowBeforeRightClickEvent
        {
            add
            {
                CreateEventBridge();
                _WindowBeforeRightClickEvent += value;
            }
            remove
            {
                _WindowBeforeRightClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_WindowBeforeDoubleClickEventHandler _WindowBeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745746.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_WindowBeforeDoubleClickEventHandler WindowBeforeDoubleClickEvent
        {
            add
            {
                CreateEventBridge();
                _WindowBeforeDoubleClickEvent += value;
            }
            remove
            {
                _WindowBeforeDoubleClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationCloseEventHandler _PresentationCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744678.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationCloseEventHandler PresentationCloseEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationCloseEvent += value;
            }
            remove
            {
                _PresentationCloseEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationSaveEventHandler _PresentationSaveEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744230.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationSaveEventHandler PresentationSaveEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationSaveEvent += value;
            }
            remove
            {
                _PresentationSaveEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationOpenEventHandler _PresentationOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744100.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationOpenEventHandler PresentationOpenEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationOpenEvent += value;
            }
            remove
            {
                _PresentationOpenEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_NewPresentationEventHandler _NewPresentationEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745073.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_NewPresentationEventHandler NewPresentationEvent
        {
            add
            {
                CreateEventBridge();
                _NewPresentationEvent += value;
            }
            remove
            {
                _NewPresentationEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationNewSlideEventHandler _PresentationNewSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746597.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationNewSlideEventHandler PresentationNewSlideEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationNewSlideEvent += value;
            }
            remove
            {
                _PresentationNewSlideEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_WindowActivateEventHandler _WindowActivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff743995.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_WindowActivateEventHandler WindowActivateEvent
        {
            add
            {
                CreateEventBridge();
                _WindowActivateEvent += value;
            }
            remove
            {
                _WindowActivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_WindowDeactivateEventHandler _WindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745519.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_WindowDeactivateEventHandler WindowDeactivateEvent
        {
            add
            {
                CreateEventBridge();
                _WindowDeactivateEvent += value;
            }
            remove
            {
                _WindowDeactivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_SlideShowBeginEventHandler _SlideShowBeginEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746741.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_SlideShowBeginEventHandler SlideShowBeginEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowBeginEvent += value;
            }
            remove
            {
                _SlideShowBeginEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_SlideShowNextBuildEventHandler _SlideShowNextBuildEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745070.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_SlideShowNextBuildEventHandler SlideShowNextBuildEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowNextBuildEvent += value;
            }
            remove
            {
                _SlideShowNextBuildEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_SlideShowNextSlideEventHandler _SlideShowNextSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745863.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_SlideShowNextSlideEventHandler SlideShowNextSlideEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowNextSlideEvent += value;
            }
            remove
            {
                _SlideShowNextSlideEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_SlideShowEndEventHandler _SlideShowEndEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746536.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_SlideShowEndEventHandler SlideShowEndEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowEndEvent += value;
            }
            remove
            {
                _SlideShowEndEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 9,10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationPrintEventHandler _PresentationPrintEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 9 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744696.aspx </remarks>
        [SupportByVersion("PowerPoint", 9, 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationPrintEventHandler PresentationPrintEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationPrintEvent += value;
            }
            remove
            {
                _PresentationPrintEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 10,11,12,14,15,16
        /// </summary>
        private event Application_SlideSelectionChangedEventHandler _SlideSelectionChangedEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745869.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public event Application_SlideSelectionChangedEventHandler SlideSelectionChangedEvent
        {
            add
            {
                CreateEventBridge();
                _SlideSelectionChangedEvent += value;
            }
            remove
            {
                _SlideSelectionChangedEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 10,11,12,14,15,16
        /// </summary>
        private event Application_ColorSchemeChangedEventHandler _ColorSchemeChangedEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745549.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public event Application_ColorSchemeChangedEventHandler ColorSchemeChangedEvent
        {
            add
            {
                CreateEventBridge();
                _ColorSchemeChangedEvent += value;
            }
            remove
            {
                _ColorSchemeChangedEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 10,11,12,14,15,16
        /// </summary>
        private event Application_PresentationBeforeSaveEventHandler _PresentationBeforeSaveEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744682.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public event Application_PresentationBeforeSaveEventHandler PresentationBeforeSaveEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationBeforeSaveEvent += value;
            }
            remove
            {
                _PresentationBeforeSaveEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 10,11,12,14,15,16
        /// </summary>
        private event Application_SlideShowNextClickEventHandler _SlideShowNextClickEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 10 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745682.aspx </remarks>
        [SupportByVersion("PowerPoint", 10, 11, 12, 14, 15, 16)]
        public event Application_SlideShowNextClickEventHandler SlideShowNextClickEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowNextClickEvent += value;
            }
            remove
            {
                _SlideShowNextClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 11,12,14,15,16
        /// </summary>
        private event Application_AfterNewPresentationEventHandler _AfterNewPresentationEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746421.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        public event Application_AfterNewPresentationEventHandler AfterNewPresentationEvent
        {
            add
            {
                CreateEventBridge();
                _AfterNewPresentationEvent += value;
            }
            remove
            {
                _AfterNewPresentationEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 11,12,14,15,16
        /// </summary>
        private event Application_AfterPresentationOpenEventHandler _AfterPresentationOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744659.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        public event Application_AfterPresentationOpenEventHandler AfterPresentationOpenEvent
        {
            add
            {
                CreateEventBridge();
                _AfterPresentationOpenEvent += value;
            }
            remove
            {
                _AfterPresentationOpenEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 11,12,14,15,16
        /// </summary>
        private event Application_PresentationSyncEventHandler _PresentationSyncEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 11 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744576.aspx </remarks>
        [SupportByVersion("PowerPoint", 11, 12, 14, 15, 16)]
        public event Application_PresentationSyncEventHandler PresentationSyncEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationSyncEvent += value;
            }
            remove
            {
                _PresentationSyncEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 12,14,15,16
        /// </summary>
        private event Application_SlideShowOnNextEventHandler _SlideShowOnNextEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746469.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public event Application_SlideShowOnNextEventHandler SlideShowOnNextEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowOnNextEvent += value;
            }
            remove
            {
                _SlideShowOnNextEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 12,14,15,16
        /// </summary>
        private event Application_SlideShowOnPreviousEventHandler _SlideShowOnPreviousEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 12 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744749.aspx </remarks>
        [SupportByVersion("PowerPoint", 12, 14, 15, 16)]
        public event Application_SlideShowOnPreviousEventHandler SlideShowOnPreviousEvent
        {
            add
            {
                CreateEventBridge();
                _SlideShowOnPreviousEvent += value;
            }
            remove
            {
                _SlideShowOnPreviousEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_PresentationBeforeCloseEventHandler _PresentationBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745567.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_PresentationBeforeCloseEventHandler PresentationBeforeCloseEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationBeforeCloseEvent += value;
            }
            remove
            {
                _PresentationBeforeCloseEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_ProtectedViewWindowOpenEventHandler _ProtectedViewWindowOpenEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745081.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_ProtectedViewWindowOpenEventHandler ProtectedViewWindowOpenEvent
        {
            add
            {
                CreateEventBridge();
                _ProtectedViewWindowOpenEvent += value;
            }
            remove
            {
                _ProtectedViewWindowOpenEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_ProtectedViewWindowBeforeEditEventHandler _ProtectedViewWindowBeforeEditEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745575.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_ProtectedViewWindowBeforeEditEventHandler ProtectedViewWindowBeforeEditEvent
        {
            add
            {
                CreateEventBridge();
                _ProtectedViewWindowBeforeEditEvent += value;
            }
            remove
            {
                _ProtectedViewWindowBeforeEditEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_ProtectedViewWindowBeforeCloseEventHandler _ProtectedViewWindowBeforeCloseEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746497.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_ProtectedViewWindowBeforeCloseEventHandler ProtectedViewWindowBeforeCloseEvent
        {
            add
            {
                CreateEventBridge();
                _ProtectedViewWindowBeforeCloseEvent += value;
            }
            remove
            {
                _ProtectedViewWindowBeforeCloseEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_ProtectedViewWindowActivateEventHandler _ProtectedViewWindowActivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744591.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_ProtectedViewWindowActivateEventHandler ProtectedViewWindowActivateEvent
        {
            add
            {
                CreateEventBridge();
                _ProtectedViewWindowActivateEvent += value;
            }
            remove
            {
                _ProtectedViewWindowActivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_ProtectedViewWindowDeactivateEventHandler _ProtectedViewWindowDeactivateEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746253.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_ProtectedViewWindowDeactivateEventHandler ProtectedViewWindowDeactivateEvent
        {
            add
            {
                CreateEventBridge();
                _ProtectedViewWindowDeactivateEvent += value;
            }
            remove
            {
                _ProtectedViewWindowDeactivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 14,15,16
        /// </summary>
        private event Application_PresentationCloseFinalEventHandler _PresentationCloseFinalEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 14 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff744781.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public event Application_PresentationCloseFinalEventHandler PresentationCloseFinalEvent
        {
            add
            {
                CreateEventBridge();
                _PresentationCloseFinalEvent += value;
            }
            remove
            {
                _PresentationCloseFinalEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 15, 16
        /// </summary>
        private event Application_AfterDragDropOnSlideEventHandler _AfterDragDropOnSlideEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227644.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public event Application_AfterDragDropOnSlideEventHandler AfterDragDropOnSlideEvent
        {
            add
            {
                CreateEventBridge();
                _AfterDragDropOnSlideEvent += value;
            }
            remove
            {
                _AfterDragDropOnSlideEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint, 15, 16
        /// </summary>
        private event Application_AfterShapeSizeChangeEventHandler _AfterShapeSizeChangeEvent;

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227375.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public event Application_AfterShapeSizeChangeEventHandler AfterShapeSizeChangeEvent
        {
            add
            {
                CreateEventBridge();
                _AfterShapeSizeChangeEvent += value;
            }
            remove
            {
                _AfterShapeSizeChangeEvent -= value;
            }
        }

        #endregion

        #region IAutomaticQuit

        /// <summary>
        /// Determines Quit method want be called while disposing if NetOffice.Settings.EnableAutomaticQuit is true.
        /// Default is true when instance has no parent object and its not a cloned instance, otherwise false.
        /// </summary>
        bool IAutomaticQuit.Enabled
        {

            get
            {
                return _callQuitInDispose;
            }
            set
            {
                _callQuitInDispose = value;
            }
        }

        #endregion

        #pragma warning restore
    }
}
