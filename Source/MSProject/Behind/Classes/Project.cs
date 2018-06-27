using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.Behind
{
	/// <summary>
	/// CoClass Project 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920664(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.MSProjectApi.EventContracts._EProjectDoc))]
    public class Project : _IProjectDoc, NetOffice.MSProjectApi.Project
    {
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private NetOffice.MSProjectApi.Behind.EventContracts._EProjectDoc_SinkHelper __EProjectDoc_SinkHelper;
	
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
                    _contractType = typeof(NetOffice.MSProjectApi.Project);
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
                    _type = typeof(Project);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctor

		/// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>		
		public Project():base()
		{
			
		}	

		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_OpenEventHandler _OpenEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_OpenEventHandler OpenEvent
		{
			add
			{
				CreateEventBridge();
				_OpenEvent += value;
			}
			remove
			{
				_OpenEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_BeforeCloseEventHandler _BeforeCloseEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_BeforeCloseEventHandler BeforeCloseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeCloseEvent += value;
			}
			remove
			{
				_BeforeCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_BeforeSaveEventHandler _BeforeSaveEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_BeforeSaveEventHandler BeforeSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSaveEvent += value;
			}
			remove
			{
				_BeforeSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_BeforePrintEventHandler _BeforePrintEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_BeforePrintEventHandler BeforePrintEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePrintEvent += value;
			}
			remove
			{
				_BeforePrintEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_CalculateEventHandler _CalculateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_CalculateEventHandler CalculateEvent
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_ChangeEventHandler _ChangeEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_ChangeEventHandler ChangeEvent
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_ActivateEventHandler _ActivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_ActivateEventHandler ActivateEvent
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
		/// SupportByVersion MSProject, 11,12,14
		/// </summary>
		private event Project_DeactivateEventHandler _DeactivateEvent;

		/// <summary>
		/// SupportByVersion MSProject 11 12 14
		/// </summary>
		[SupportByVersion("MSProject", 11,12,14)]
		public virtual event Project_DeactivateEventHandler DeactivateEvent
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

		#endregion
       
	    #region IEventBinding
        
		/// <summary>
        /// Creates active sink helper
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual void CreateEventBridge()
        {
			if(false == Factory.Settings.EnableEvents)
				return;
	
			if (null != _connectPoint)
				return;
	
            if (null == _activeSinkId)
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.MSProjectApi.Behind.EventContracts._EProjectDoc_SinkHelper.Id);


			if(NetOffice.MSProjectApi.Behind.EventContracts._EProjectDoc_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				__EProjectDoc_SinkHelper = new NetOffice.MSProjectApi.Behind.EventContracts._EProjectDoc_SinkHelper(this, _connectPoint);
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
			if( null != __EProjectDoc_SinkHelper)
			{
				__EProjectDoc_SinkHelper.Dispose();
				__EProjectDoc_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}
