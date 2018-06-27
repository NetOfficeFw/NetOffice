using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
    /// <summary>
    /// CoClass Recordset 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.ADODBApi.EventContracts.RecordsetEvents))]
    public class Recordset : _Recordset, NetOffice.ADODBApi.Recordset
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private NetOffice.ADODBApi.Behind.EventContracts.RecordsetEvents_SinkHelper _recordsetEvents_SinkHelper;

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
                    _contractType = typeof(NetOffice.ADODBApi.Recordset);
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
                    _type = typeof(Recordset);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>		
        public Recordset() : base()
        {

        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_WillChangeFieldEventHandler _WillChangeFieldEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_WillChangeFieldEventHandler WillChangeFieldEvent
        {
            add
            {
                CreateEventBridge();
                _WillChangeFieldEvent += value;
            }
            remove
            {
                _WillChangeFieldEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_FieldChangeCompleteEventHandler _FieldChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_FieldChangeCompleteEventHandler FieldChangeCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _FieldChangeCompleteEvent += value;
            }
            remove
            {
                _FieldChangeCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_WillChangeRecordEventHandler _WillChangeRecordEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_WillChangeRecordEventHandler WillChangeRecordEvent
        {
            add
            {
                CreateEventBridge();
                _WillChangeRecordEvent += value;
            }
            remove
            {
                _WillChangeRecordEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_RecordChangeCompleteEventHandler _RecordChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_RecordChangeCompleteEventHandler RecordChangeCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _RecordChangeCompleteEvent += value;
            }
            remove
            {
                _RecordChangeCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_WillChangeRecordsetEventHandler _WillChangeRecordsetEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_WillChangeRecordsetEventHandler WillChangeRecordsetEvent
        {
            add
            {
                CreateEventBridge();
                _WillChangeRecordsetEvent += value;
            }
            remove
            {
                _WillChangeRecordsetEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_RecordsetChangeCompleteEventHandler _RecordsetChangeCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_RecordsetChangeCompleteEventHandler RecordsetChangeCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _RecordsetChangeCompleteEvent += value;
            }
            remove
            {
                _RecordsetChangeCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_WillMoveEventHandler _WillMoveEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_WillMoveEventHandler WillMoveEvent
        {
            add
            {
                CreateEventBridge();
                _WillMoveEvent += value;
            }
            remove
            {
                _WillMoveEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_MoveCompleteEventHandler _MoveCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_MoveCompleteEventHandler MoveCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _MoveCompleteEvent += value;
            }
            remove
            {
                _MoveCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_EndOfRecordsetEventHandler _EndOfRecordsetEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_EndOfRecordsetEventHandler EndOfRecordsetEvent
        {
            add
            {
                CreateEventBridge();
                _EndOfRecordsetEvent += value;
            }
            remove
            {
                _EndOfRecordsetEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_FetchProgressEventHandler _FetchProgressEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_FetchProgressEventHandler FetchProgressEvent
        {
            add
            {
                CreateEventBridge();
                _FetchProgressEvent += value;
            }
            remove
            {
                _FetchProgressEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Recordset_FetchCompleteEventHandler _FetchCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Recordset_FetchCompleteEventHandler FetchCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _FetchCompleteEvent += value;
            }
            remove
            {
                _FetchCompleteEvent -= value;
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
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.ADODBApi.Behind.EventContracts.RecordsetEvents_SinkHelper.Id);


            if (NetOffice.ADODBApi.Behind.EventContracts.RecordsetEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _recordsetEvents_SinkHelper = new NetOffice.ADODBApi.Behind.EventContracts.RecordsetEvents_SinkHelper(this, _connectPoint);
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
            if (null != _recordsetEvents_SinkHelper)
            {
                _recordsetEvents_SinkHelper.Dispose();
                _recordsetEvents_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #pragma warning restore
    }
}
