using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.MSFormsApi
{

	#region Delegates

	#pragma warning disable
	public delegate void Frame_AddControlEventHandler(NetOffice.MSFormsApi.Control Control);
	public delegate void Frame_BeforeDragOverEventHandler(NetOffice.MSFormsApi.ReturnBoolean Cancel, NetOffice.MSFormsApi.Control Control, NetOffice.MSFormsApi.DataObject Data, Single X, Single Y, NetOffice.MSFormsApi.Enums.fmDragState State, NetOffice.MSFormsApi.ReturnEffect Effect, Int16 Shift);
	public delegate void Frame_BeforeDropOrPasteEventHandler(NetOffice.MSFormsApi.ReturnBoolean Cancel, NetOffice.MSFormsApi.Control Control, NetOffice.MSFormsApi.Enums.fmAction Action, NetOffice.MSFormsApi.DataObject Data, Single X, Single Y, NetOffice.MSFormsApi.ReturnEffect Effect, Int16 Shift);
	public delegate void Frame_ClickEventHandler();
	public delegate void Frame_DblClickEventHandler(NetOffice.MSFormsApi.ReturnBoolean Cancel);
	public delegate void Frame_ErrorEventHandler(Int16 Number, NetOffice.MSFormsApi.ReturnString Description, Int32 SCode, string Source, string HelpFile, Int32 HelpContext, NetOffice.MSFormsApi.ReturnBoolean CancelDisplay);
	public delegate void Frame_KeyDownEventHandler(NetOffice.MSFormsApi.ReturnInteger KeyCode, Int16 Shift);
	public delegate void Frame_KeyPressEventHandler(NetOffice.MSFormsApi.ReturnInteger KeyAscii);
	public delegate void Frame_KeyUpEventHandler(NetOffice.MSFormsApi.ReturnInteger KeyCode, Int16 Shift);
	public delegate void Frame_LayoutEventHandler();
	public delegate void Frame_MouseDownEventHandler(Int16 Button, Int16 Shift, Single X, Single Y);
	public delegate void Frame_MouseMoveEventHandler(Int16 Button, Int16 Shift, Single X, Single Y);
	public delegate void Frame_MouseUpEventHandler(Int16 Button, Int16 Shift, Single X, Single Y);
	public delegate void Frame_RemoveControlEventHandler(NetOffice.MSFormsApi.Control Control);
	public delegate void Frame_ScrollEventHandler(NetOffice.MSFormsApi.Enums.fmScrollAction ActionX, NetOffice.MSFormsApi.Enums.fmScrollAction ActionY, Single RequestDx, Single RequestDy, NetOffice.MSFormsApi.ReturnSingle ActualDx, NetOffice.MSFormsApi.ReturnSingle ActualDy);
	public delegate void Frame_ZoomEventHandler(ref Int16 Percent);
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass Frame 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class Frame : IOptionFrame,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		OptionFrameEvents_SinkHelper _optionFrameEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(Frame);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Frame(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Frame(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Frame(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Frame(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Frame(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of Frame 
        ///</summary>		
		public Frame():base("MSForms.Frame")
		{
			
		}
		
		///<summary>
        ///creates a new instance of Frame
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public Frame(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running MSForms.Frame objects from the running object table(ROT)
        /// </summary>
        /// <returns>an MSForms.Frame array</returns>
		public static NetOffice.MSFormsApi.Frame[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("MSForms","Frame");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.MSFormsApi.Frame> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.MSFormsApi.Frame>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.MSFormsApi.Frame(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running MSForms.Frame object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an MSForms.Frame object or null</returns>
		public static NetOffice.MSFormsApi.Frame GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("MSForms","Frame", false);
			if(null != proxy)
				return new NetOffice.MSFormsApi.Frame(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running MSForms.Frame object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an MSForms.Frame object or null</returns>
		public static NetOffice.MSFormsApi.Frame GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("MSForms","Frame", throwOnError);
			if(null != proxy)
				return new NetOffice.MSFormsApi.Frame(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_AddControlEventHandler _AddControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_AddControlEventHandler AddControlEvent
		{
			add
			{
				CreateEventBridge();
				_AddControlEvent += value;
			}
			remove
			{
				_AddControlEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_BeforeDragOverEventHandler _BeforeDragOverEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_BeforeDragOverEventHandler BeforeDragOverEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDragOverEvent += value;
			}
			remove
			{
				_BeforeDragOverEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_BeforeDropOrPasteEventHandler _BeforeDropOrPasteEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_BeforeDropOrPasteEventHandler BeforeDropOrPasteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDropOrPasteEvent += value;
			}
			remove
			{
				_BeforeDropOrPasteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_ClickEventHandler _ClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_ClickEventHandler ClickEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_DblClickEventHandler _DblClickEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_DblClickEventHandler DblClickEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_ErrorEventHandler _ErrorEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_ErrorEventHandler ErrorEvent
		{
			add
			{
				CreateEventBridge();
				_ErrorEvent += value;
			}
			remove
			{
				_ErrorEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_LayoutEventHandler _LayoutEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_LayoutEventHandler LayoutEvent
		{
			add
			{
				CreateEventBridge();
				_LayoutEvent += value;
			}
			remove
			{
				_LayoutEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_RemoveControlEventHandler _RemoveControlEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_RemoveControlEventHandler RemoveControlEvent
		{
			add
			{
				CreateEventBridge();
				_RemoveControlEvent += value;
			}
			remove
			{
				_RemoveControlEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_ScrollEventHandler _ScrollEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_ScrollEventHandler ScrollEvent
		{
			add
			{
				CreateEventBridge();
				_ScrollEvent += value;
			}
			remove
			{
				_ScrollEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms, 2
		/// </summary>
		private event Frame_ZoomEventHandler _ZoomEvent;

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public event Frame_ZoomEventHandler ZoomEvent
		{
			add
			{
				CreateEventBridge();
				_ZoomEvent += value;
			}
			remove
			{
				_ZoomEvent -= value;
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
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, OptionFrameEvents_SinkHelper.Id);


			if(OptionFrameEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_optionFrameEvents_SinkHelper = new OptionFrameEvents_SinkHelper(this, _connectPoint);
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
                        Factory.Console.WriteException(exception);
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
			if( null != _optionFrameEvents_SinkHelper)
			{
				_optionFrameEvents_SinkHelper.Dispose();
				_optionFrameEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}