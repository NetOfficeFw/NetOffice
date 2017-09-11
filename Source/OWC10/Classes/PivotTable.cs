using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	#region Delegates

	#pragma warning disable
	public delegate void PivotTable_SelectionChangeEventHandler();
	public delegate void PivotTable_ViewChangeEventHandler(NetOffice.OWC10Api.Enums.PivotViewReasonEnum reason);
	public delegate void PivotTable_DataChangeEventHandler(NetOffice.OWC10Api.Enums.PivotDataReasonEnum reason);
	public delegate void PivotTable_PivotTableChangeEventHandler(NetOffice.OWC10Api.Enums.PivotTableReasonEnum reason);
	public delegate void PivotTable_BeforeQueryEventHandler();
	public delegate void PivotTable_QueryEventHandler();
	public delegate void PivotTable_OnConnectEventHandler();
	public delegate void PivotTable_OnDisconnectEventHandler();
	public delegate void PivotTable_MouseDownEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseMoveEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseUpEventHandler(Int32 button, Int32 shift, Int32 x, Int32 y);
	public delegate void PivotTable_MouseWheelEventHandler(bool page, Int32 count);
	public delegate void PivotTable_ClickEventHandler();
	public delegate void PivotTable_DblClickEventHandler();
	public delegate void PivotTable_CommandEnabledEventHandler(object command, NetOffice.OWC10Api.ByRef enabled);
	public delegate void PivotTable_CommandCheckedEventHandler(object command, NetOffice.OWC10Api.ByRef Checked);
	public delegate void PivotTable_CommandTipTextEventHandler(object command, NetOffice.OWC10Api.ByRef caption);
	public delegate void PivotTable_CommandBeforeExecuteEventHandler(object command, NetOffice.OWC10Api.ByRef Cancel);
	public delegate void PivotTable_CommandExecuteEventHandler(object command, bool succeeded);
	public delegate void PivotTable_KeyDownEventHandler(Int32 keyCode, Int32 shift);
	public delegate void PivotTable_KeyUpEventHandler(Int32 keyCode, Int32 shift);
	public delegate void PivotTable_KeyPressEventHandler(Int32 keyAscii);
	public delegate void PivotTable_BeforeKeyDownEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeKeyUpEventHandler(Int32 keyCode, Int32 shift, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeKeyPressEventHandler(Int32 keyAscii, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_BeforeContextMenuEventHandler(Int32 x, Int32 y, NetOffice.OWC10Api.ByRef Menu, NetOffice.OWC10Api.ByRef cancel);
	public delegate void PivotTable_StartEditEventHandler(ICOMObject selection, ICOMObject activeObject, NetOffice.OWC10Api.ByRef initialValue, NetOffice.OWC10Api.ByRef arrowMode, NetOffice.OWC10Api.ByRef caretPosition, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void PivotTable_EndEditEventHandler(bool accept, NetOffice.OWC10Api.ByRef finalValue, NetOffice.OWC10Api.ByRef cancel, NetOffice.OWC10Api.ByRef errorDescription);
	public delegate void PivotTable_BeforeScreenTipEventHandler(NetOffice.OWC10Api.ByRef screenTipText, ICOMObject sourceObject);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass PivotTable 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.IPivotControlEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.IPivotControlEvents))]
    public class PivotTable : IPivotControl, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.IPivotControlEvents_SinkHelper _iPivotControlEvents_SinkHelper;
	
		#endregion

		#region Type Information

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
                    _type = typeof(PivotTable);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotTable(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotTable(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotTable(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotTable(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotTable(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of PivotTable 
        /// </summary>		
		public PivotTable():base("OWC10.PivotTable")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of PivotTable
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public PivotTable(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_SelectionChangeEventHandler _SelectionChangeEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_SelectionChangeEventHandler SelectionChangeEvent
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
		private event PivotTable_ViewChangeEventHandler _ViewChangeEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_ViewChangeEventHandler ViewChangeEvent
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

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_DataChangeEventHandler _DataChangeEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_DataChangeEventHandler DataChangeEvent
		{
			add
			{
				CreateEventBridge();
				_DataChangeEvent += value;
			}
			remove
			{
				_DataChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_PivotTableChangeEventHandler _PivotTableChangeEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_PivotTableChangeEventHandler PivotTableChangeEvent
		{
			add
			{
				CreateEventBridge();
				_PivotTableChangeEvent += value;
			}
			remove
			{
				_PivotTableChangeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_BeforeQueryEventHandler _BeforeQueryEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeQueryEventHandler BeforeQueryEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeQueryEvent += value;
			}
			remove
			{
				_BeforeQueryEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_QueryEventHandler _QueryEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_QueryEventHandler QueryEvent
		{
			add
			{
				CreateEventBridge();
				_QueryEvent += value;
			}
			remove
			{
				_QueryEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_OnConnectEventHandler _OnConnectEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_OnConnectEventHandler OnConnectEvent
		{
			add
			{
				CreateEventBridge();
				_OnConnectEvent += value;
			}
			remove
			{
				_OnConnectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_OnDisconnectEventHandler _OnDisconnectEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_OnDisconnectEventHandler OnDisconnectEvent
		{
			add
			{
				CreateEventBridge();
				_OnDisconnectEvent += value;
			}
			remove
			{
				_OnDisconnectEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_MouseDownEventHandler MouseDownEvent
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
		private event PivotTable_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion OWC10, 1
		/// </summary>
		private event PivotTable_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_MouseUpEventHandler MouseUpEvent
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
		private event PivotTable_MouseWheelEventHandler _MouseWheelEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_MouseWheelEventHandler MouseWheelEvent
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
		private event PivotTable_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_ClickEventHandler ClickEvent
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
		private event PivotTable_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_DblClickEventHandler DblClickEvent
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
		private event PivotTable_CommandEnabledEventHandler _CommandEnabledEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_CommandEnabledEventHandler CommandEnabledEvent
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
		private event PivotTable_CommandCheckedEventHandler _CommandCheckedEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_CommandCheckedEventHandler CommandCheckedEvent
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
		private event PivotTable_CommandTipTextEventHandler _CommandTipTextEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_CommandTipTextEventHandler CommandTipTextEvent
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
		private event PivotTable_CommandBeforeExecuteEventHandler _CommandBeforeExecuteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_CommandBeforeExecuteEventHandler CommandBeforeExecuteEvent
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
		private event PivotTable_CommandExecuteEventHandler _CommandExecuteEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_CommandExecuteEventHandler CommandExecuteEvent
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
		private event PivotTable_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_KeyDownEventHandler KeyDownEvent
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
		private event PivotTable_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_KeyUpEventHandler KeyUpEvent
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
		private event PivotTable_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_KeyPressEventHandler KeyPressEvent
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
		private event PivotTable_BeforeKeyDownEventHandler _BeforeKeyDownEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeKeyDownEventHandler BeforeKeyDownEvent
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
		private event PivotTable_BeforeKeyUpEventHandler _BeforeKeyUpEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeKeyUpEventHandler BeforeKeyUpEvent
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
		private event PivotTable_BeforeKeyPressEventHandler _BeforeKeyPressEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeKeyPressEventHandler BeforeKeyPressEvent
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
		private event PivotTable_BeforeContextMenuEventHandler _BeforeContextMenuEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeContextMenuEventHandler BeforeContextMenuEvent
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
		private event PivotTable_StartEditEventHandler _StartEditEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_StartEditEventHandler StartEditEvent
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
		private event PivotTable_EndEditEventHandler _EndEditEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_EndEditEventHandler EndEditEvent
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
		private event PivotTable_BeforeScreenTipEventHandler _BeforeScreenTipEvent;

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public event PivotTable_BeforeScreenTipEventHandler BeforeScreenTipEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeScreenTipEvent += value;
			}
			remove
			{
				_BeforeScreenTipEvent -= value;
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
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.IPivotControlEvents_SinkHelper.Id);


			if(Events.IPivotControlEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_iPivotControlEvents_SinkHelper = new Events.IPivotControlEvents_SinkHelper(this, _connectPoint);
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
			if( null != _iPivotControlEvents_SinkHelper)
			{
				_iPivotControlEvents_SinkHelper.Dispose();
				_iPivotControlEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

