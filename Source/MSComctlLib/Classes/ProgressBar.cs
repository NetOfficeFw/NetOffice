using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.MSComctlLibApi
{

	#region Delegates

	#pragma warning disable
	public delegate void ProgressBar_MouseDownEventHandler(Int16 Button, Int16 Shift, Int32 x, Int32 y);
	public delegate void ProgressBar_MouseMoveEventHandler(Int16 Button, Int16 Shift, Int32 x, Int32 y);
	public delegate void ProgressBar_MouseUpEventHandler(Int16 Button, Int16 Shift, Int32 x, Int32 y);
	public delegate void ProgressBar_ClickEventHandler();
	public delegate void ProgressBar_OLEStartDragEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int32 AllowedEffects);
	public delegate void ProgressBar_OLEGiveFeedbackEventHandler(ref Int32 Effect, ref bool DefaultCursors);
	public delegate void ProgressBar_OLESetDataEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int16 DataFormat);
	public delegate void ProgressBar_OLECompleteDragEventHandler(ref Int32 Effect);
	public delegate void ProgressBar_OLEDragOverEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int32 Effect, ref Int16 Button, ref Int16 Shift, ref Single x, ref Single y, ref Int16 State);
	public delegate void ProgressBar_OLEDragDropEventHandler(ref NetOffice.MSComctlLibApi.DataObject Data, ref Int32 Effect, ref Int16 Button, ref Int16 Shift, ref Single x, ref Single y);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass ProgressBar 
	/// SupportByVersion MSComctlLib, 6.0
	///</summary>
	[SupportByVersionAttribute("MSComctlLib", 6.0)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class ProgressBar : IProgressBar,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		IProgressBarEvents_SinkHelper _iProgressBarEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(ProgressBar);
                    
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
		public ProgressBar(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProgressBar(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ProgressBar(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// creates a new instance of ProgressBar 
        /// </summary>		
		public ProgressBar():base("MSComctlLib.ProgressBar")
		{
			
		}
		
		/// <summary>
        /// creates a new instance of ProgressBar
        /// </summary>
        /// <param name="progId">registered ProgID</param>
		public ProgressBar(string progId):base(progId)
		{
			
		}

		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_ClickEventHandler ClickEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLEStartDragEventHandler _OLEStartDragEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLEStartDragEventHandler OLEStartDragEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLEGiveFeedbackEventHandler _OLEGiveFeedbackEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLEGiveFeedbackEventHandler OLEGiveFeedbackEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLESetDataEventHandler _OLESetDataEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLESetDataEventHandler OLESetDataEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLECompleteDragEventHandler _OLECompleteDragEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLECompleteDragEventHandler OLECompleteDragEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLEDragOverEventHandler _OLEDragOverEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLEDragOverEventHandler OLEDragOverEvent
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
		/// SupportByVersion MSComctlLib, 6.0
		/// </summary>
		private event ProgressBar_OLEDragDropEventHandler _OLEDragDropEvent;

		/// <summary>
		/// SupportByVersion MSComctlLib 6.0
		/// </summary>
		[SupportByVersion("MSComctlLib", 6.0)]
		public event ProgressBar_OLEDragDropEventHandler OLEDragDropEvent
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
       
	    #region IEventBinding Member
        
		/// <summary>
        /// creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void CreateEventBridge()
        {
			if(false == NetOffice.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, IProgressBarEvents_SinkHelper.Id);


			if(IProgressBarEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_iProgressBarEvents_SinkHelper = new IProgressBarEvents_SinkHelper(this, _connectPoint);
				return;
			} 
        }

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool EventBridgeInitialized
        {
            get 
            {
                return (null != _connectPoint);
            }
        }
        
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public bool HasEventRecipients()       
        {
			if(null == _thisType)
				_thisType = this.GetType();
					
			foreach (NetRuntimeSystem.Reflection.EventInfo item in _thisType.GetEvents())
			{
				MulticastDelegate eventDelegate = (MulticastDelegate) _thisType.GetType().GetField(item.Name, 
																			NetRuntimeSystem.Reflection.BindingFlags.NonPublic |
																			NetRuntimeSystem.Reflection.BindingFlags.Instance).GetValue(this);
					
				if( (null != eventDelegate) && (eventDelegate.GetInvocationList().Length > 0) )
					return false;
			}
				
			return false;
        }
        
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public Delegate[] GetEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates;
            }
            else
                return new Delegate[0];
        }

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int GetCountOfEventRecipients(string eventName)
        {
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                return delegates.Length;
            }
            else
                return 0;
        }

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public int RaiseCustomEvent(string eventName, ref object[] paramsArray)
		{
			if(null == _thisType)
				_thisType = this.GetType();
             
            MulticastDelegate eventDelegate = (MulticastDelegate)_thisType.GetField(
                                                "_" + eventName + "Event",
                                                NetRuntimeSystem.Reflection.BindingFlags.Instance |
                                                NetRuntimeSystem.Reflection.BindingFlags.NonPublic).GetValue(this);

            if (null != eventDelegate)
            {
                Delegate[] delegates = eventDelegate.GetInvocationList();
                foreach (var item in delegates)
                {
                    try
                    {
                        item.Method.Invoke(item.Target, paramsArray);
                    }
                    catch (NetRuntimeSystem.Exception exception)
                    {
                        DebugConsole.WriteException(exception);
                    }
                }
                return delegates.Length;
            }
            else
                return 0;
		}

        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public void DisposeEventBridge()
        {
			if( null != _iProgressBarEvents_SinkHelper)
			{
				_iProgressBarEvents_SinkHelper.Dispose();
				_iProgressBarEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}