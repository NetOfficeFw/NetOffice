using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSComctlLibApi.Behind
{
    /// <summary>
    /// CoClass TreeView 
    /// SupportByVersion MSComctlLib, 6
    /// </summary>
    [SupportByVersion("MSComctlLib", 6)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSComctlLibApi.EventContracts.ITreeViewEvents))]
    public class TreeView : ITreeView, NetOffice.MSComctlLibApi.TreeView
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private NetOffice.MSComctlLibApi.Behind.EventContracts.ITreeViewEvents_SinkHelper _iTreeViewEvents_SinkHelper;

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
                    _contractType = typeof(NetOffice.MSComctlLibApi.TreeView);
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
                    _type = typeof(TreeView);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>		
        public TreeView() : base()
        {

        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_BeforeLabelEditEventHandler _BeforeLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_BeforeLabelEditEventHandler BeforeLabelEditEvent
        {
            add
            {
                CreateEventBridge();
                _BeforeLabelEditEvent += value;
            }
            remove
            {
                _BeforeLabelEditEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_AfterLabelEditEventHandler _AfterLabelEditEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_AfterLabelEditEventHandler AfterLabelEditEvent
        {
            add
            {
                CreateEventBridge();
                _AfterLabelEditEvent += value;
            }
            remove
            {
                _AfterLabelEditEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_CollapseEventHandler _CollapseEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_CollapseEventHandler CollapseEvent
        {
            add
            {
                CreateEventBridge();
                _CollapseEvent += value;
            }
            remove
            {
                _CollapseEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_ExpandEventHandler _ExpandEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_ExpandEventHandler ExpandEvent
        {
            add
            {
                CreateEventBridge();
                _ExpandEvent += value;
            }
            remove
            {
                _ExpandEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_NodeClickEventHandler _NodeClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_NodeClickEventHandler NodeClickEvent
        {
            add
            {
                CreateEventBridge();
                _NodeClickEvent += value;
            }
            remove
            {
                _NodeClickEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_KeyDownEventHandler _KeyDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_KeyDownEventHandler KeyDownEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_KeyUpEventHandler _KeyUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_KeyUpEventHandler KeyUpEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_KeyPressEventHandler _KeyPressEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_KeyPressEventHandler KeyPressEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_MouseDownEventHandler _MouseDownEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_MouseDownEventHandler MouseDownEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_MouseMoveEventHandler _MouseMoveEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_MouseMoveEventHandler MouseMoveEvent
        {
            add
            {
                CreateEventBridge();
                _MouseMoveEvent += value;
            }
            remove
            {
                _MouseMoveEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_MouseUpEventHandler _MouseUpEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_MouseUpEventHandler MouseUpEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_ClickEventHandler _ClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_ClickEventHandler ClickEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_DblClickEventHandler _DblClickEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_DblClickEventHandler DblClickEvent
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
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_NodeCheckEventHandler _NodeCheckEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_NodeCheckEventHandler NodeCheckEvent
        {
            add
            {
                CreateEventBridge();
                _NodeCheckEvent += value;
            }
            remove
            {
                _NodeCheckEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLEStartDragEventHandler _OLEStartDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLEStartDragEventHandler OLEStartDragEvent
        {
            add
            {
                CreateEventBridge();
                _OLEStartDragEvent += value;
            }
            remove
            {
                _OLEStartDragEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLEGiveFeedbackEventHandler _OLEGiveFeedbackEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent
        {
            add
            {
                CreateEventBridge();
                _OLEGiveFeedbackEvent += value;
            }
            remove
            {
                _OLEGiveFeedbackEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLESetDataEventHandler _OLESetDataEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLESetDataEventHandler OLESetDataEvent
        {
            add
            {
                CreateEventBridge();
                _OLESetDataEvent += value;
            }
            remove
            {
                _OLESetDataEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLECompleteDragEventHandler _OLECompleteDragEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLECompleteDragEventHandler OLECompleteDragEvent
        {
            add
            {
                CreateEventBridge();
                _OLECompleteDragEvent += value;
            }
            remove
            {
                _OLECompleteDragEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLEDragOverEventHandler _OLEDragOverEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLEDragOverEventHandler OLEDragOverEvent
        {
            add
            {
                CreateEventBridge();
                _OLEDragOverEvent += value;
            }
            remove
            {
                _OLEDragOverEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion MSComctlLib, 6
        /// </summary>
        private event TreeView_OLEDragDropEventHandler _OLEDragDropEvent;

        /// <summary>
        /// SupportByVersion MSComctlLib 6
        /// </summary>
        [SupportByVersion("MSComctlLib", 6)]
        public event TreeView_OLEDragDropEventHandler OLEDragDropEvent
        {
            add
            {
                CreateEventBridge();
                _OLEDragDropEvent += value;
            }
            remove
            {
                _OLEDragDropEvent -= value;
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
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.MSComctlLibApi.Behind.EventContracts.ITreeViewEvents_SinkHelper.Id);

            if (NetOffice.MSComctlLibApi.Behind.EventContracts.ITreeViewEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _iTreeViewEvents_SinkHelper = new NetOffice.MSComctlLibApi.Behind.EventContracts.ITreeViewEvents_SinkHelper(this, _connectPoint);
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
            if (null != _iTreeViewEvents_SinkHelper)
            {
                _iTreeViewEvents_SinkHelper.Dispose();
                _iTreeViewEvents_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #pragma warning restore
    }
}
