using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice;
namespace NetOffice.OfficeApi
{

	#region Delegates

	#pragma warning disable
	public delegate void MsoEnvelope_EnvelopeShowEventHandler();
	public delegate void MsoEnvelope_EnvelopeHideEventHandler();
	#pragma warning restore

	#endregion

	///<summary>
	/// CoClass MsoEnvelope 
	/// SupportByVersion Office, 10,11,12,14,15
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff862112.aspx
	///</summary>
	[SupportByVersionAttribute("Office", 10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsCoClass)]
	public class MsoEnvelope : IMsoEnvelopeVB,IEventBinding
	{
		#pragma warning disable
		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
		private NetRuntimeSystem.Type _thisType;
		IMsoEnvelopeVBEvents_SinkHelper _iMsoEnvelopeVBEvents_SinkHelper;
	
		#endregion

		#region Type Information

        private static Type _type;
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(MsoEnvelope);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MsoEnvelope(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public MsoEnvelope(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoEnvelope(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoEnvelope(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public MsoEnvelope(COMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		///<summary>
        ///creates a new instance of MsoEnvelope 
        ///</summary>		
		public MsoEnvelope():base("Office.MsoEnvelope")
		{
			
		}
		
		///<summary>
        ///creates a new instance of MsoEnvelope
        ///</summary>
        ///<param name="progId">registered ProgID</param>
		public MsoEnvelope(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods

		/// <summary>
        /// returns all running Office.MsoEnvelope objects from the running object table(ROT)
        /// </summary>
        /// <returns>an Office.MsoEnvelope array</returns>
		public static NetOffice.OfficeApi.MsoEnvelope[] GetActiveInstances()
		{		
			NetRuntimeSystem.Collections.Generic.List<object> proxyList = NetOffice.RunningObjectTable.GetActiveProxiesFromROT("Office","MsoEnvelope");
			NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.MsoEnvelope> resultList = new NetRuntimeSystem.Collections.Generic.List<NetOffice.OfficeApi.MsoEnvelope>();
			foreach(object proxy in proxyList)
				resultList.Add( new NetOffice.OfficeApi.MsoEnvelope(null, proxy) );
			return resultList.ToArray();
		}

		/// <summary>
        /// returns a running Office.MsoEnvelope object from the running object table(ROT). the method takes the first element from the table
        /// </summary>
        /// <returns>an Office.MsoEnvelope object or null</returns>
		public static NetOffice.OfficeApi.MsoEnvelope GetActiveInstance()
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","MsoEnvelope", false);
			if(null != proxy)
				return new NetOffice.OfficeApi.MsoEnvelope(null, proxy);
			else
				return null;
		}

		/// <summary>
        /// returns a running Office.MsoEnvelope object from the running object table(ROT).  the method takes the first element from the table
        /// </summary>
	    /// <param name="throwOnError">throw an exception if no object was found</param>
        /// <returns>an Office.MsoEnvelope object or null</returns>
		public static NetOffice.OfficeApi.MsoEnvelope GetActiveInstance(bool throwOnError)
		{
			object proxy = NetOffice.RunningObjectTable.GetActiveProxyFromROT("Office","MsoEnvelope", throwOnError);
			if(null != proxy)
				return new NetOffice.OfficeApi.MsoEnvelope(null, proxy);
			else
				return null;
		}
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Office, 10,11,12,14,15
		/// </summary>
		private event MsoEnvelope_EnvelopeShowEventHandler _EnvelopeShowEvent;

		/// <summary>
		/// SupportByVersion Office 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff861098.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15)]
		public event MsoEnvelope_EnvelopeShowEventHandler EnvelopeShowEvent
		{
			add
			{
				CreateEventBridge();
				_EnvelopeShowEvent += value;
			}
			remove
			{
				_EnvelopeShowEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Office, 10,11,12,14,15
		/// </summary>
		private event MsoEnvelope_EnvelopeHideEventHandler _EnvelopeHideEvent;

		/// <summary>
		/// SupportByVersion Office 10 11 12 14 15
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff860254.aspx </remarks>
		[SupportByVersion("Office", 10,11,12,14,15)]
		public event MsoEnvelope_EnvelopeHideEventHandler EnvelopeHideEvent
		{
			add
			{
				CreateEventBridge();
				_EnvelopeHideEvent += value;
			}
			remove
			{
				_EnvelopeHideEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, IMsoEnvelopeVBEvents_SinkHelper.Id);


			if(IMsoEnvelopeVBEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_iMsoEnvelopeVBEvents_SinkHelper = new IMsoEnvelopeVBEvents_SinkHelper(this, _connectPoint);
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
			if( null != _iMsoEnvelopeVBEvents_SinkHelper)
			{
				_iMsoEnvelopeVBEvents_SinkHelper.Dispose();
				_iMsoEnvelopeVBEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}