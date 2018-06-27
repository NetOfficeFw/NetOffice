using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ADODBApi;

namespace NetOffice.ADODBApi.Behind
{
    /// <summary>
    /// CoClass Connection 
    /// SupportByVersion ADODB, 2.1,2.5
    /// </summary>
    [SupportByVersion("ADODB", 2.1, 2.5)]
    [EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.ADODBApi.EventContracts.ConnectionEvents))]
    public class Connection : _Connection, NetOffice.ADODBApi.Connection
    {
        #pragma warning disable

        #region Fields

        private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
        private string _activeSinkId;
        private static Type _type;
        private NetOffice.ADODBApi.Behind.EventContracts.ConnectionEvents_SinkHelper _connectionEvents_SinkHelper;

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
                    _contractType = typeof(NetOffice.ADODBApi.Connection);
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
                    _type = typeof(Connection);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Strub Ctor, not intended to use
        /// </summary>		
        public Connection() : base()
        {

        }

        #endregion

        #region Events

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_InfoMessageEventHandler _InfoMessageEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_InfoMessageEventHandler InfoMessageEvent
        {
            add
            {
                CreateEventBridge();
                _InfoMessageEvent += value;
            }
            remove
            {
                _InfoMessageEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_BeginTransCompleteEventHandler _BeginTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_BeginTransCompleteEventHandler BeginTransCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _BeginTransCompleteEvent += value;
            }
            remove
            {
                _BeginTransCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_CommitTransCompleteEventHandler _CommitTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_CommitTransCompleteEventHandler CommitTransCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _CommitTransCompleteEvent += value;
            }
            remove
            {
                _CommitTransCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_RollbackTransCompleteEventHandler _RollbackTransCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_RollbackTransCompleteEventHandler RollbackTransCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _RollbackTransCompleteEvent += value;
            }
            remove
            {
                _RollbackTransCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_WillExecuteEventHandler _WillExecuteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_WillExecuteEventHandler WillExecuteEvent
        {
            add
            {
                CreateEventBridge();
                _WillExecuteEvent += value;
            }
            remove
            {
                _WillExecuteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_ExecuteCompleteEventHandler _ExecuteCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_ExecuteCompleteEventHandler ExecuteCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _ExecuteCompleteEvent += value;
            }
            remove
            {
                _ExecuteCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_WillConnectEventHandler _WillConnectEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_WillConnectEventHandler WillConnectEvent
        {
            add
            {
                CreateEventBridge();
                _WillConnectEvent += value;
            }
            remove
            {
                _WillConnectEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_ConnectCompleteEventHandler _ConnectCompleteEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_ConnectCompleteEventHandler ConnectCompleteEvent
        {
            add
            {
                CreateEventBridge();
                _ConnectCompleteEvent += value;
            }
            remove
            {
                _ConnectCompleteEvent -= value;
            }
        }

        /// <summary>
        /// SupportByVersion ADODB, 2.1,2.5
        /// </summary>
        private event Connection_DisconnectEventHandler _DisconnectEvent;

        /// <summary>
        /// SupportByVersion ADODB 2.1 2.5
        /// </summary>
        [SupportByVersion("ADODB", 2.1, 2.5)]
        public virtual event Connection_DisconnectEventHandler DisconnectEvent
        {
            add
            {
                CreateEventBridge();
                _DisconnectEvent += value;
            }
            remove
            {
                _DisconnectEvent -= value;
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
                _activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.ADODBApi.Behind.EventContracts.ConnectionEvents_SinkHelper.Id);


            if (NetOffice.ADODBApi.Behind.EventContracts.ConnectionEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
            {
                _connectionEvents_SinkHelper = new NetOffice.ADODBApi.Behind.EventContracts.ConnectionEvents_SinkHelper(this, _connectPoint);
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
            if (null != _connectionEvents_SinkHelper)
            {
                _connectionEvents_SinkHelper.Dispose();
                _connectionEvents_SinkHelper = null;
            }

            _connectPoint = null;
        }

        #endregion

        #pragma warning restore
    }
}
