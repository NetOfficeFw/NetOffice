using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// CoClass Worksheet
    /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194464.aspx </remarks>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.ExcelApi.EventContracts.DocEvents))]
    public class Worksheet : NetOffice.ExcelApi.Behind._Worksheet, NetOffice.ExcelApi.Worksheet
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private EventContracts.DocEvents_SinkHelper _docEvents_SinkHelper;

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
                    _contractType = typeof(NetOffice.ExcelApi.Worksheet);
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
                    _type = typeof(Worksheet);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Worksheet() : base()
        {

        }

        #endregion

        #region Static CoClass Methods

        /// <summary>
        /// Returns all running Excel.Worksheet instances from the environment/system
        /// </summary>
        /// <returns>Excel.Worksheet sequence</returns>
        public static IDisposableSequence<Application> GetActiveInstances()
        {
            return ProxyService.GetActiveInstances<Application>("Excel", "Worksheet");
        }

        /// <summary>
        /// Returns a running Excel.Worksheet instance from the environment/system
        /// </summary>
        /// <param name="throwExceptionIfNotFound">throw exception if unable to find an instance</param>
        /// <returns>Excel.Worksheet instance or null</returns>
        public static Application GetActiveInstance(bool throwExceptionIfNotFound = false)
        {
            return ProxyService.GetActiveInstance<Application>("Excel", "Worksheet", throwExceptionIfNotFound);
        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_SelectionChangeEventHandler _SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194470.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_SelectionChangeEventHandler SelectionChangeEvent
        {
            add
            {
                CreateEventBridge();
                _SelectionChangeEvent += value;
            }
            remove
            {
                _SelectionChangeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_BeforeDoubleClickEventHandler _BeforeDoubleClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff196564.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_BeforeDoubleClickEventHandler BeforeDoubleClickEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeDoubleClickEvent += value;
            }
            remove
            {
                _BeforeDoubleClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_BeforeRightClickEventHandler _BeforeRightClickEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff192993.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_BeforeRightClickEventHandler BeforeRightClickEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeRightClickEvent += value;
            }
            remove
            {
                _BeforeRightClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_ActivateEventHandler _ActivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198220.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_ActivateEventHandler ActivateEvent
        {
            add
            {
                CreateEventBridge();
                _ActivateEvent += value;
            }
            remove
            {
                _ActivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_DeactivateEventHandler _DeactivateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197183.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_DeactivateEventHandler DeactivateEvent
        {
            add
            {
                CreateEventBridge();
                _DeactivateEvent += value;
            }
            remove
            {
                _DeactivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_CalculateEventHandler _CalculateEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838823.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_CalculateEventHandler CalculateEvent
        {
            add
            {
                CreateEventBridge();
                _CalculateEvent += value;
            }
            remove
            {
                _CalculateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_ChangeEventHandler _ChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff839775.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_ChangeEventHandler ChangeEvent
        {
            add
            {
                CreateEventBridge();
                _ChangeEvent += value;
            }
            remove
            {
                _ChangeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_FollowHyperlinkEventHandler _FollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838843.aspx </remarks>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_FollowHyperlinkEventHandler FollowHyperlinkEvent
        {
            add
            {
                CreateEventBridge();
                _FollowHyperlinkEvent += value;
            }
            remove
            {
                _FollowHyperlinkEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableUpdateEventHandler _PivotTableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 10, 11, 12, 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff822105.aspx </remarks>
        [SupportByVersion("Excel", 10, 11, 12, 14, 15, 16)]
        public virtual event Worksheet_PivotTableUpdateEventHandler PivotTableUpdateEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableUpdateEvent += value;
            }
            remove
            {
                _PivotTableUpdateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableAfterValueChangeEventHandler _PivotTableAfterValueChangeEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193517.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Worksheet_PivotTableAfterValueChangeEventHandler PivotTableAfterValueChangeEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableAfterValueChangeEvent += value;
            }
            remove
            {
                _PivotTableAfterValueChangeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableBeforeAllocateChangesEventHandler _PivotTableBeforeAllocateChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff195070.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Worksheet_PivotTableBeforeAllocateChangesEventHandler PivotTableBeforeAllocateChangesEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableBeforeAllocateChangesEvent += value;
            }
            remove
            {
                _PivotTableBeforeAllocateChangesEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableBeforeCommitChangesEventHandler _PivotTableBeforeCommitChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff198138.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Worksheet_PivotTableBeforeCommitChangesEventHandler PivotTableBeforeCommitChangesEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableBeforeCommitChangesEvent += value;
            }
            remove
            {
                _PivotTableBeforeCommitChangesEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableBeforeDiscardChangesEventHandler _PivotTableBeforeDiscardChangesEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836187.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Worksheet_PivotTableBeforeDiscardChangesEventHandler PivotTableBeforeDiscardChangesEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableBeforeDiscardChangesEvent += value;
            }
            remove
            {
                _PivotTableBeforeDiscardChangesEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        private event Worksheet_PivotTableChangeSyncEventHandler _PivotTableChangeSyncEvent;

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838251.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual event Worksheet_PivotTableChangeSyncEventHandler PivotTableChangeSyncEvent
        {
            add
            {
                CreateEventBridge();
                _PivotTableChangeSyncEvent += value;
            }
            remove
            {
                _PivotTableChangeSyncEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Worksheet_LensGalleryRenderCompleteEventHandler _LensGalleryRenderCompleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj227603.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Worksheet_LensGalleryRenderCompleteEventHandler LensGalleryRenderCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _LensGalleryRenderCompleteEvent += value;
            }
            remove
            {
                _LensGalleryRenderCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Worksheet_TableUpdateEventHandler _TableUpdateEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/jj229788.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Worksheet_TableUpdateEventHandler TableUpdateEvent
        {
            add
            {
                CreateEventBridge();
                _TableUpdateEvent += value;
            }
            remove
            {
                _TableUpdateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        private event Worksheet_BeforeDeleteEventHandler _BeforeDeleteEvent;

        /// <summary>
        /// SupportByVersion Excel 15, 16
        /// </summary>
        ///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/dn448393.aspx </remarks>
        [SupportByVersion("Excel", 15, 16)]
        public virtual event Worksheet_BeforeDeleteEventHandler BeforeDeleteEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeDeleteEvent += value;
            }
            remove
            {
                _BeforeDeleteEvent -= value;
            }
        }

        #endregion

        #region IEventBinding

        /// <summary>
        /// Creates active sink helper
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void CreateEventBridge()
        {
            if (false == Factory.Settings.EnableEvents)
                return;

            if (null != _connectPoint)
                return;

            if (null == _activeSinkId)
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, EventContracts.DocEvents_SinkHelper.Id);


            if (EventContracts.DocEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _docEvents_SinkHelper = new EventContracts.DocEvents_SinkHelper(this, _connectPoint);
                return;
            }
        }

        /// <summary>
        /// The instance use currently an event listener
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool EventBridgeInitialized
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
        public virtual bool HasEventRecipients()
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType);
        }

        /// <summary>
        /// Instance has one or more event recipients
        /// </summary>
        /// <param name="eventName">name of the event</param>
        /// <returns></returns>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.HasEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Target methods from its actual event recipients
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Delegate[] GetEventRecipients(string eventName)
        {
            return NetOffice.Events.CoClassEventReflector.GetEventRecipients(this, LateBindingApiWrapperType, eventName);
        }

        /// <summary>
        /// Returns the current count of event recipients
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual int GetCountOfEventRecipients(string eventName)
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
        public virtual int RaiseCustomEvent(string eventName, ref object[] paramsArray)
        {
            return NetOffice.Events.CoClassEventReflector.RaiseCustomEvent(this, LateBindingApiWrapperType, eventName, ref paramsArray);
        }
        /// <summary>
        /// Stop listening events for the instance
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void DisposeEventBridge()
        {
            if (null != _docEvents_SinkHelper)
            {
                _docEvents_SinkHelper.Dispose();
                _docEvents_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #pragma warning restore
    }
}
