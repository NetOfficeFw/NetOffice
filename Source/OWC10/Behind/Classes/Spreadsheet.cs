using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// CoClass Spreadsheet 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.OWC10Api.EventContracts.ISpreadsheetEventSink))]
    public class Spreadsheet : ISpreadsheet, IEventBinding
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private NetOffice.OWC10Api.Behind.EventContracts.ISpreadsheetEventSink_SinkHelper _iSpreadsheetEventSink_SinkHelper;

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
                    _contractType = typeof(NetOffice.OWC10Api.Spreadsheet);
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
                    _type = typeof(Spreadsheet);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>		
        public Spreadsheet() : base()
        {

        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_BeforeContextMenuEventHandler _BeforeContextMenuEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_BeforeContextMenuEventHandler BeforeContextMenuEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeContextMenuEvent += value;
            }
            remove
            {
                _BeforeContextMenuEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_BeforeKeyDownEventHandler _BeforeKeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_BeforeKeyDownEventHandler BeforeKeyDownEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeKeyDownEvent += value;
            }
            remove
            {
                _BeforeKeyDownEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_BeforeKeyPressEventHandler _BeforeKeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_BeforeKeyPressEventHandler BeforeKeyPressEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeKeyPressEvent += value;
            }
            remove
            {
                _BeforeKeyPressEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_BeforeKeyUpEventHandler _BeforeKeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_BeforeKeyUpEventHandler BeforeKeyUpEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeKeyUpEvent += value;
            }
            remove
            {
                _BeforeKeyUpEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_ClickEventHandler _ClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_ClickEventHandler ClickEvent
        {
            add
            {
                CreateEventBridge();
                _ClickEvent += value;
            }
            remove
            {
                _ClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_CommandEnabledEventHandler _CommandEnabledEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_CommandEnabledEventHandler CommandEnabledEvent
        {
            add
            {
                CreateEventBridge();
                _CommandEnabledEvent += value;
            }
            remove
            {
                _CommandEnabledEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_CommandCheckedEventHandler _CommandCheckedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_CommandCheckedEventHandler CommandCheckedEvent
        {
            add
            {
                CreateEventBridge();
                _CommandCheckedEvent += value;
            }
            remove
            {
                _CommandCheckedEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_CommandTipTextEventHandler _CommandTipTextEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_CommandTipTextEventHandler CommandTipTextEvent
        {
            add
            {
                CreateEventBridge();
                _CommandTipTextEvent += value;
            }
            remove
            {
                _CommandTipTextEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_CommandBeforeExecuteEventHandler _CommandBeforeExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent
        {
            add
            {
                CreateEventBridge();
                _CommandBeforeExecuteEvent += value;
            }
            remove
            {
                _CommandBeforeExecuteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_CommandExecuteEventHandler _CommandExecuteEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_CommandExecuteEventHandler CommandExecuteEvent
        {
            add
            {
                CreateEventBridge();
                _CommandExecuteEvent += value;
            }
            remove
            {
                _CommandExecuteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_DblClickEventHandler _DblClickEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_DblClickEventHandler DblClickEvent
        {
            add
            {
                CreateEventBridge();
                _DblClickEvent += value;
            }
            remove
            {
                _DblClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_EndEditEventHandler _EndEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_EndEditEventHandler EndEditEvent
        {
            add
            {
                CreateEventBridge();
                _EndEditEvent += value;
            }
            remove
            {
                _EndEditEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_InitializeEventHandler _InitializeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_InitializeEventHandler InitializeEvent
        {
            add
            {
                CreateEventBridge();
                _InitializeEvent += value;
            }
            remove
            {
                _InitializeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_KeyDownEventHandler _KeyDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_KeyDownEventHandler KeyDownEvent
        {
            add
            {
                CreateEventBridge();
                _KeyDownEvent += value;
            }
            remove
            {
                _KeyDownEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_KeyPressEventHandler _KeyPressEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_KeyPressEventHandler KeyPressEvent
        {
            add
            {
                CreateEventBridge();
                _KeyPressEvent += value;
            }
            remove
            {
                _KeyPressEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_KeyUpEventHandler _KeyUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_KeyUpEventHandler KeyUpEvent
        {
            add
            {
                CreateEventBridge();
                _KeyUpEvent += value;
            }
            remove
            {
                _KeyUpEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_LoadCompletedEventHandler _LoadCompletedEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_LoadCompletedEventHandler LoadCompletedEvent
        {
            add
            {
                CreateEventBridge();
                _LoadCompletedEvent += value;
            }
            remove
            {
                _LoadCompletedEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_MouseDownEventHandler _MouseDownEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_MouseDownEventHandler MouseDownEvent
        {
            add
            {
                CreateEventBridge();
                _MouseDownEvent += value;
            }
            remove
            {
                _MouseDownEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_MouseOutEventHandler _MouseOutEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_MouseOutEventHandler MouseOutEvent
        {
            add
            {
                CreateEventBridge();
                _MouseOutEvent += value;
            }
            remove
            {
                _MouseOutEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_MouseOverEventHandler _MouseOverEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_MouseOverEventHandler MouseOverEvent
        {
            add
            {
                CreateEventBridge();
                _MouseOverEvent += value;
            }
            remove
            {
                _MouseOverEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_MouseUpEventHandler _MouseUpEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_MouseUpEventHandler MouseUpEvent
        {
            add
            {
                CreateEventBridge();
                _MouseUpEvent += value;
            }
            remove
            {
                _MouseUpEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_MouseWheelEventHandler _MouseWheelEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_MouseWheelEventHandler MouseWheelEvent
        {
            add
            {
                CreateEventBridge();
                _MouseWheelEvent += value;
            }
            remove
            {
                _MouseWheelEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SelectionChangeEventHandler _SelectionChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SelectionChangeEventHandler SelectionChangeEvent
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
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SelectionChangingEventHandler _SelectionChangingEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SelectionChangingEventHandler SelectionChangingEvent
        {
            add
            {
                CreateEventBridge();
                _SelectionChangingEvent += value;
            }
            remove
            {
                _SelectionChangingEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SheetActivateEventHandler _SheetActivateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SheetActivateEventHandler SheetActivateEvent
        {
            add
            {
                CreateEventBridge();
                _SheetActivateEvent += value;
            }
            remove
            {
                _SheetActivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SheetCalculateEventHandler _SheetCalculateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SheetCalculateEventHandler SheetCalculateEvent
        {
            add
            {
                CreateEventBridge();
                _SheetCalculateEvent += value;
            }
            remove
            {
                _SheetCalculateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SheetChangeEventHandler _SheetChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SheetChangeEventHandler SheetChangeEvent
        {
            add
            {
                CreateEventBridge();
                _SheetChangeEvent += value;
            }
            remove
            {
                _SheetChangeEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SheetDeactivateEventHandler _SheetDeactivateEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SheetDeactivateEventHandler SheetDeactivateEvent
        {
            add
            {
                CreateEventBridge();
                _SheetDeactivateEvent += value;
            }
            remove
            {
                _SheetDeactivateEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_SheetFollowHyperlinkEventHandler _SheetFollowHyperlinkEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_SheetFollowHyperlinkEventHandler SheetFollowHyperlinkEvent
        {
            add
            {
                CreateEventBridge();
                _SheetFollowHyperlinkEvent += value;
            }
            remove
            {
                _SheetFollowHyperlinkEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_StartEditEventHandler _StartEditEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_StartEditEventHandler StartEditEvent
        {
            add
            {
                CreateEventBridge();
                _StartEditEvent += value;
            }
            remove
            {
                _StartEditEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion OWC10, 1
        /// </summary>
        private event Spreadsheet_ViewChangeEventHandler _ViewChangeEvent;

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual event Spreadsheet_ViewChangeEventHandler ViewChangeEvent
        {
            add
            {
                CreateEventBridge();
                _ViewChangeEvent += value;
            }
            remove
            {
                _ViewChangeEvent -= value;
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
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.OWC10Api.Behind.EventContracts.ISpreadsheetEventSink_SinkHelper.Id);


            if (NetOffice.OWC10Api.Behind.EventContracts.ISpreadsheetEventSink_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _iSpreadsheetEventSink_SinkHelper = new NetOffice.OWC10Api.Behind.EventContracts.ISpreadsheetEventSink_SinkHelper(this, _connectPoint);
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
            if (null != _iSpreadsheetEventSink_SinkHelper)
            {
                _iSpreadsheetEventSink_SinkHelper.Dispose();
                _iSpreadsheetEventSink_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #pragma warning restore
    }
}
