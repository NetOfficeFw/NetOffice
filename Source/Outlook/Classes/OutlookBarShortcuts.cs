﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	#region Delegates

	#pragma warning disable
	public delegate void OutlookBarShortcuts_ShortcutAddEventHandler(NetOffice.OutlookApi.OutlookBarShortcut newShortcut);
	public delegate void OutlookBarShortcuts_BeforeShortcutAddEventHandler(ref bool cancel);
	public delegate void OutlookBarShortcuts_BeforeShortcutRemoveEventHandler(NetOffice.OutlookApi.OutlookBarShortcut shortcut, ref bool cancel);
	#pragma warning restore

	#endregion

	/// <summary>
	/// CoClass OutlookBarShortcuts 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OutlookBarShortcuts"/> </remarks>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
	[EventSink(typeof(Events.OutlookBarShortcutsEvents_SinkHelper))]
    [ComEventInterface(typeof(Events.OutlookBarShortcutsEvents))]
    public class OutlookBarShortcuts : _OutlookBarShortcuts, IEventBinding
	{
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private Events.OutlookBarShortcutsEvents_SinkHelper _outlookBarShortcutsEvents_SinkHelper;
	
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
                    _type = typeof(OutlookBarShortcuts);
                return _type;
            }
        }
        
        #endregion
        		
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OutlookBarShortcuts(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public OutlookBarShortcuts(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
			
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarShortcuts(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{
			
		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarShortcuts(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
			
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OutlookBarShortcuts(ICOMObject replacedObject) : base(replacedObject)
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of OutlookBarShortcuts 
        /// </summary>		
		public OutlookBarShortcuts():base("Outlook.OutlookBarShortcuts")
		{
			
		}
		
		/// <summary>
        /// Creates a new instance of OutlookBarShortcuts
        /// </summary>
        ///<param name="progId">registered ProgID</param>
		public OutlookBarShortcuts(string progId):base(progId)
		{
			
		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event OutlookBarShortcuts_ShortcutAddEventHandler _ShortcutAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OutlookBarShortcuts.ShortcutAdd"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event OutlookBarShortcuts_ShortcutAddEventHandler ShortcutAddEvent
		{
			add
			{
				CreateEventBridge();
				_ShortcutAddEvent += value;
			}
			remove
			{
				_ShortcutAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event OutlookBarShortcuts_BeforeShortcutAddEventHandler _BeforeShortcutAddEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OutlookBarShortcuts.BeforeShortcutAdd"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event OutlookBarShortcuts_BeforeShortcutAddEventHandler BeforeShortcutAddEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShortcutAddEvent += value;
			}
			remove
			{
				_BeforeShortcutAddEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook, 9,10,11,12,14,15,16
		/// </summary>
		private event OutlookBarShortcuts_BeforeShortcutRemoveEventHandler _BeforeShortcutRemoveEvent;

		/// <summary>
		/// SupportByVersion Outlook 9 10 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.OutlookBarShortcuts.BeforeShortcutRemove"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public event OutlookBarShortcuts_BeforeShortcutRemoveEventHandler BeforeShortcutRemoveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShortcutRemoveEvent += value;
			}
			remove
			{
				_BeforeShortcutRemoveEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, Events.OutlookBarShortcutsEvents_SinkHelper.Id);


			if(Events.OutlookBarShortcutsEvents_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_outlookBarShortcutsEvents_SinkHelper = new Events.OutlookBarShortcutsEvents_SinkHelper(this, _connectPoint);
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
			if( null != _outlookBarShortcutsEvents_SinkHelper)
			{
				_outlookBarShortcutsEvents_SinkHelper.Dispose();
				_outlookBarShortcutsEvents_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

