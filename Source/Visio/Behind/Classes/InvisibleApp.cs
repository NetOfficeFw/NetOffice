using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// CoClass InvisibleApp 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769303(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EApplication))]
    public class InvisibleApp : IVInvisibleApp, NetOffice.VisioApi.InvisibleApp
    {
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private NetOffice.VisioApi.Behind.EventContracts.EApplication_SinkHelper _eApplication_SinkHelper;
	
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
                    _contractType = typeof(NetOffice.VisioApi.InvisibleApp);
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
                    _type = typeof(InvisibleApp);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public InvisibleApp() : base()
		{

		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AppActivatedEventHandler _AppActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767393(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AppActivatedEventHandler AppActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppActivatedEvent += value;
			}
			remove
			{
				_AppActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AppDeactivatedEventHandler _AppDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765522(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AppDeactivatedEventHandler AppDeactivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppDeactivatedEvent += value;
			}
			remove
			{
				_AppDeactivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AppObjActivatedEventHandler _AppObjActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768447(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AppObjActivatedEventHandler AppObjActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppObjActivatedEvent += value;
			}
			remove
			{
				_AppObjActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AppObjDeactivatedEventHandler _AppObjDeactivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765372(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AppObjDeactivatedEventHandler AppObjDeactivatedEvent
		{
			add
			{
				CreateEventBridge();
				_AppObjDeactivatedEvent += value;
			}
			remove
			{
				_AppObjDeactivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeQuitEventHandler _BeforeQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767922(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeQuitEventHandler BeforeQuitEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeQuitEvent += value;
			}
			remove
			{
				_BeforeQuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeModalEventHandler _BeforeModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767572(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeModalEventHandler BeforeModalEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeModalEvent += value;
			}
			remove
			{
				_BeforeModalEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AfterModalEventHandler _AfterModalEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766355(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AfterModalEventHandler AfterModalEvent
		{
			add
			{
				CreateEventBridge();
				_AfterModalEvent += value;
			}
			remove
			{
				_AfterModalEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_WindowOpenedEventHandler _WindowOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767410(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_WindowOpenedEventHandler WindowOpenedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowOpenedEvent += value;
			}
			remove
			{
				_WindowOpenedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_SelectionChangedEventHandler _SelectionChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766811(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_SelectionChangedEventHandler SelectionChangedEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionChangedEvent += value;
			}
			remove
			{
				_SelectionChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeWindowClosedEventHandler _BeforeWindowClosedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768040(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeWindowClosedEventHandler BeforeWindowClosedEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowClosedEvent += value;
			}
			remove
			{
				_BeforeWindowClosedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_WindowActivatedEventHandler _WindowActivatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767376(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_WindowActivatedEventHandler WindowActivatedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowActivatedEvent += value;
			}
			remove
			{
				_WindowActivatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeWindowSelDeleteEventHandler _BeforeWindowSelDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766033(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeWindowSelDeleteEventHandler BeforeWindowSelDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowSelDeleteEvent += value;
			}
			remove
			{
				_BeforeWindowSelDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeWindowPageTurnEventHandler _BeforeWindowPageTurnEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767063(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeWindowPageTurnEventHandler BeforeWindowPageTurnEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeWindowPageTurnEvent += value;
			}
			remove
			{
				_BeforeWindowPageTurnEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_WindowTurnedToPageEventHandler _WindowTurnedToPageEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767662(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_WindowTurnedToPageEventHandler WindowTurnedToPageEvent
		{
			add
			{
				CreateEventBridge();
				_WindowTurnedToPageEvent += value;
			}
			remove
			{
				_WindowTurnedToPageEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentOpenedEventHandler _DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766381(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentOpenedEventHandler DocumentOpenedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentOpenedEvent += value;
			}
			remove
			{
				_DocumentOpenedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentCreatedEventHandler _DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767356(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentCreatedEventHandler DocumentCreatedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentCreatedEvent += value;
			}
			remove
			{
				_DocumentCreatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentSavedEventHandler _DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768368(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentSavedEventHandler DocumentSavedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentSavedEvent += value;
			}
			remove
			{
				_DocumentSavedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentSavedAsEventHandler _DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769114(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentSavedAsEventHandler DocumentSavedAsEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentSavedAsEvent += value;
			}
			remove
			{
				_DocumentSavedAsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentChangedEventHandler _DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768511(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentChangedEventHandler DocumentChangedEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentChangedEvent += value;
			}
			remove
			{
				_DocumentChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeDocumentCloseEventHandler _BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766731(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentCloseEvent += value;
			}
			remove
			{
				_BeforeDocumentCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_StyleAddedEventHandler _StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767954(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_StyleAddedEventHandler StyleAddedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleAddedEvent += value;
			}
			remove
			{
				_StyleAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_StyleChangedEventHandler _StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767281(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_StyleChangedEventHandler StyleChangedEvent
		{
			add
			{
				CreateEventBridge();
				_StyleChangedEvent += value;
			}
			remove
			{
				_StyleChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeStyleDeleteEventHandler _BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765113(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeStyleDeleteEvent += value;
			}
			remove
			{
				_BeforeStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MasterAddedEventHandler _MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766331(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MasterAddedEventHandler MasterAddedEvent
		{
			add
			{
				CreateEventBridge();
				_MasterAddedEvent += value;
			}
			remove
			{
				_MasterAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MasterChangedEventHandler _MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768066(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MasterChangedEventHandler MasterChangedEvent
		{
			add
			{
				CreateEventBridge();
				_MasterChangedEvent += value;
			}
			remove
			{
				_MasterChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeMasterDeleteEventHandler _BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767032(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeMasterDeleteEvent += value;
			}
			remove
			{
				_BeforeMasterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_PageAddedEventHandler _PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768709(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_PageAddedEventHandler PageAddedEvent
		{
			add
			{
				CreateEventBridge();
				_PageAddedEvent += value;
			}
			remove
			{
				_PageAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_PageChangedEventHandler _PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768785(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_PageChangedEventHandler PageChangedEvent
		{
			add
			{
				CreateEventBridge();
				_PageChangedEvent += value;
			}
			remove
			{
				_PageChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforePageDeleteEventHandler _BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768579(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforePageDeleteEventHandler BeforePageDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforePageDeleteEvent += value;
			}
			remove
			{
				_BeforePageDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeAddedEventHandler _ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767730(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ShapeAddedEventHandler ShapeAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeAddedEvent += value;
			}
			remove
			{
				_ShapeAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeSelectionDeleteEventHandler _BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767690(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSelectionDeleteEvent += value;
			}
			remove
			{
				_BeforeSelectionDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeChangedEventHandler _ShapeChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767622(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ShapeChangedEventHandler ShapeChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeChangedEvent += value;
			}
			remove
			{
				_ShapeChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_SelectionAddedEventHandler _SelectionAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769104(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_SelectionAddedEventHandler SelectionAddedEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionAddedEvent += value;
			}
			remove
			{
				_SelectionAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeShapeDeleteEventHandler _BeforeShapeDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767037(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeShapeDeleteEventHandler BeforeShapeDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShapeDeleteEvent += value;
			}
			remove
			{
				_BeforeShapeDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_TextChangedEventHandler _TextChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766913(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_TextChangedEventHandler TextChangedEvent
		{
			add
			{
				CreateEventBridge();
				_TextChangedEvent += value;
			}
			remove
			{
				_TextChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_CellChangedEventHandler _CellChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766877(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_CellChangedEventHandler CellChangedEvent
		{
			add
			{
				CreateEventBridge();
				_CellChangedEvent += value;
			}
			remove
			{
				_CellChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MarkerEventEventHandler _MarkerEventEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765651(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MarkerEventEventHandler MarkerEventEvent
		{
			add
			{
				CreateEventBridge();
				_MarkerEventEvent += value;
			}
			remove
			{
				_MarkerEventEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_NoEventsPendingEventHandler _NoEventsPendingEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766723(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_NoEventsPendingEventHandler NoEventsPendingEvent
		{
			add
			{
				CreateEventBridge();
				_NoEventsPendingEvent += value;
			}
			remove
			{
				_NoEventsPendingEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_VisioIsIdleEventHandler _VisioIsIdleEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766985(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_VisioIsIdleEventHandler VisioIsIdleEvent
		{
			add
			{
				CreateEventBridge();
				_VisioIsIdleEvent += value;
			}
			remove
			{
				_VisioIsIdleEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MustFlushScopeBeginningEventHandler _MustFlushScopeBeginningEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768314(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MustFlushScopeBeginningEventHandler MustFlushScopeBeginningEvent
		{
			add
			{
				CreateEventBridge();
				_MustFlushScopeBeginningEvent += value;
			}
			remove
			{
				_MustFlushScopeBeginningEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MustFlushScopeEndedEventHandler _MustFlushScopeEndedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768618(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MustFlushScopeEndedEventHandler MustFlushScopeEndedEvent
		{
			add
			{
				CreateEventBridge();
				_MustFlushScopeEndedEvent += value;
			}
			remove
			{
				_MustFlushScopeEndedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_RunModeEnteredEventHandler _RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766960(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_RunModeEnteredEventHandler RunModeEnteredEvent
		{
			add
			{
				CreateEventBridge();
				_RunModeEnteredEvent += value;
			}
			remove
			{
				_RunModeEnteredEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DesignModeEnteredEventHandler _DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768667(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DesignModeEnteredEventHandler DesignModeEnteredEvent
		{
			add
			{
				CreateEventBridge();
				_DesignModeEnteredEvent += value;
			}
			remove
			{
				_DesignModeEnteredEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeDocumentSaveEventHandler _BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768909(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentSaveEvent += value;
			}
			remove
			{
				_BeforeDocumentSaveEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeDocumentSaveAsEventHandler _BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767685(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDocumentSaveAsEvent += value;
			}
			remove
			{
				_BeforeDocumentSaveAsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_FormulaChangedEventHandler _FormulaChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765269(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_FormulaChangedEventHandler FormulaChangedEvent
		{
			add
			{
				CreateEventBridge();
				_FormulaChangedEvent += value;
			}
			remove
			{
				_FormulaChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ConnectionsAddedEventHandler _ConnectionsAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766595(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ConnectionsAddedEventHandler ConnectionsAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectionsAddedEvent += value;
			}
			remove
			{
				_ConnectionsAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ConnectionsDeletedEventHandler _ConnectionsDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767251(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ConnectionsDeletedEventHandler ConnectionsDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ConnectionsDeletedEvent += value;
			}
			remove
			{
				_ConnectionsDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_EnterScopeEventHandler _EnterScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766336(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_EnterScopeEventHandler EnterScopeEvent
		{
			add
			{
				CreateEventBridge();
				_EnterScopeEvent += value;
			}
			remove
			{
				_EnterScopeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ExitScopeEventHandler _ExitScopeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768144(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ExitScopeEventHandler ExitScopeEvent
		{
			add
			{
				CreateEventBridge();
				_ExitScopeEvent += value;
			}
			remove
			{
				_ExitScopeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelQuitEventHandler _QueryCancelQuitEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768147(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelQuitEventHandler QueryCancelQuitEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelQuitEvent += value;
			}
			remove
			{
				_QueryCancelQuitEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QuitCanceledEventHandler _QuitCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766219(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QuitCanceledEventHandler QuitCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_QuitCanceledEvent += value;
			}
			remove
			{
				_QuitCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_WindowChangedEventHandler _WindowChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768984(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_WindowChangedEventHandler WindowChangedEvent
		{
			add
			{
				CreateEventBridge();
				_WindowChangedEvent += value;
			}
			remove
			{
				_WindowChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ViewChangedEventHandler _ViewChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766826(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ViewChangedEventHandler ViewChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ViewChangedEvent += value;
			}
			remove
			{
				_ViewChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelWindowCloseEventHandler _QueryCancelWindowCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767013(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelWindowCloseEventHandler QueryCancelWindowCloseEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelWindowCloseEvent += value;
			}
			remove
			{
				_QueryCancelWindowCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_WindowCloseCanceledEventHandler _WindowCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766190(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_WindowCloseCanceledEventHandler WindowCloseCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_WindowCloseCanceledEvent += value;
			}
			remove
			{
				_WindowCloseCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelDocumentCloseEventHandler _QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766898(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelDocumentCloseEvent += value;
			}
			remove
			{
				_QueryCancelDocumentCloseEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_DocumentCloseCanceledEventHandler _DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765949(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_DocumentCloseCanceledEvent += value;
			}
			remove
			{
				_DocumentCloseCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelStyleDeleteEventHandler _QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767280(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelStyleDeleteEvent += value;
			}
			remove
			{
				_QueryCancelStyleDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_StyleDeleteCanceledEventHandler _StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768716(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_StyleDeleteCanceledEvent += value;
			}
			remove
			{
				_StyleDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelMasterDeleteEventHandler _QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768817(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelMasterDeleteEvent += value;
			}
			remove
			{
				_QueryCancelMasterDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MasterDeleteCanceledEventHandler _MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767706(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_MasterDeleteCanceledEvent += value;
			}
			remove
			{
				_MasterDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelPageDeleteEventHandler _QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768454(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelPageDeleteEvent += value;
			}
			remove
			{
				_QueryCancelPageDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_PageDeleteCanceledEventHandler _PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765889(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_PageDeleteCanceledEventHandler PageDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_PageDeleteCanceledEvent += value;
			}
			remove
			{
				_PageDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeParentChangedEventHandler _ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766348(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ShapeParentChangedEventHandler ShapeParentChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeParentChangedEvent += value;
			}
			remove
			{
				_ShapeParentChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeShapeTextEditEventHandler _BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766839(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeShapeTextEditEvent += value;
			}
			remove
			{
				_BeforeShapeTextEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeExitedTextEditEventHandler _ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766390(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeExitedTextEditEvent += value;
			}
			remove
			{
				_ShapeExitedTextEditEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelSelectionDeleteEventHandler _QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768063(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSelectionDeleteEvent += value;
			}
			remove
			{
				_QueryCancelSelectionDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_SelectionDeleteCanceledEventHandler _SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769183(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SelectionDeleteCanceledEvent += value;
			}
			remove
			{
				_SelectionDeleteCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelUngroupEventHandler _QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767906(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelUngroupEventHandler QueryCancelUngroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelUngroupEvent += value;
			}
			remove
			{
				_QueryCancelUngroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_UngroupCanceledEventHandler _UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766810(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_UngroupCanceledEventHandler UngroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_UngroupCanceledEvent += value;
			}
			remove
			{
				_UngroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelConvertToGroupEventHandler _QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765068(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelConvertToGroupEvent += value;
			}
			remove
			{
				_QueryCancelConvertToGroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_ConvertToGroupCanceledEventHandler _ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765686(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_ConvertToGroupCanceledEvent += value;
			}
			remove
			{
				_ConvertToGroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelSuspendEventHandler _QueryCancelSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766235(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelSuspendEventHandler QueryCancelSuspendEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSuspendEvent += value;
			}
			remove
			{
				_QueryCancelSuspendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_SuspendCanceledEventHandler _SuspendCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766499(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_SuspendCanceledEventHandler SuspendCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SuspendCanceledEvent += value;
			}
			remove
			{
				_SuspendCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeSuspendEventHandler _BeforeSuspendEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769072(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_BeforeSuspendEventHandler BeforeSuspendEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSuspendEvent += value;
			}
			remove
			{
				_BeforeSuspendEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_AfterResumeEventHandler _AfterResumeEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765491(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_AfterResumeEventHandler AfterResumeEvent
		{
			add
			{
				CreateEventBridge();
				_AfterResumeEvent += value;
			}
			remove
			{
				_AfterResumeEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_OnKeystrokeMessageForAddonEventHandler _OnKeystrokeMessageForAddonEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767010(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_OnKeystrokeMessageForAddonEventHandler OnKeystrokeMessageForAddonEvent
		{
			add
			{
				CreateEventBridge();
				_OnKeystrokeMessageForAddonEvent += value;
			}
			remove
			{
				_OnKeystrokeMessageForAddonEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MouseDownEventHandler _MouseDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768379(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MouseDownEventHandler MouseDownEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MouseMoveEventHandler _MouseMoveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767115(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MouseMoveEventHandler MouseMoveEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_MouseUpEventHandler _MouseUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767397(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_MouseUpEventHandler MouseUpEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_KeyDownEventHandler _KeyDownEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767556(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_KeyDownEventHandler KeyDownEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_KeyPressEventHandler _KeyPressEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768495(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_KeyPressEventHandler KeyPressEvent
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
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event InvisibleApp_KeyUpEventHandler _KeyUpEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766229(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event InvisibleApp_KeyUpEventHandler KeyUpEvent
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
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelSuspendEventsEventHandler _QueryCancelSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765927(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelSuspendEventsEventHandler QueryCancelSuspendEventsEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelSuspendEventsEvent += value;
			}
			remove
			{
				_QueryCancelSuspendEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_SuspendEventsCanceledEventHandler _SuspendEventsCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765481(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_SuspendEventsCanceledEventHandler SuspendEventsCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_SuspendEventsCanceledEvent += value;
			}
			remove
			{
				_SuspendEventsCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeSuspendEventsEventHandler _BeforeSuspendEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766566(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_BeforeSuspendEventsEventHandler BeforeSuspendEventsEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeSuspendEventsEvent += value;
			}
			remove
			{
				_BeforeSuspendEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_AfterResumeEventsEventHandler _AfterResumeEventsEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765853(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_AfterResumeEventsEventHandler AfterResumeEventsEvent
		{
			add
			{
				CreateEventBridge();
				_AfterResumeEventsEvent += value;
			}
			remove
			{
				_AfterResumeEventsEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_QueryCancelGroupEventHandler _QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765459(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_QueryCancelGroupEventHandler QueryCancelGroupEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelGroupEvent += value;
			}
			remove
			{
				_QueryCancelGroupEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_GroupCanceledEventHandler _GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765402(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_GroupCanceledEventHandler GroupCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_GroupCanceledEvent += value;
			}
			remove
			{
				_GroupCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeDataGraphicChangedEventHandler _ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765850(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeDataGraphicChangedEvent += value;
			}
			remove
			{
				_ShapeDataGraphicChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_BeforeDataRecordsetDeleteEventHandler _BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765236(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeDataRecordsetDeleteEvent += value;
			}
			remove
			{
				_BeforeDataRecordsetDeleteEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_DataRecordsetChangedEventHandler _DataRecordsetChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768548(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_DataRecordsetChangedEventHandler DataRecordsetChangedEvent
		{
			add
			{
				CreateEventBridge();
				_DataRecordsetChangedEvent += value;
			}
			remove
			{
				_DataRecordsetChangedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_DataRecordsetAddedEventHandler _DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767031(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_DataRecordsetAddedEventHandler DataRecordsetAddedEvent
		{
			add
			{
				CreateEventBridge();
				_DataRecordsetAddedEvent += value;
			}
			remove
			{
				_DataRecordsetAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeLinkAddedEventHandler _ShapeLinkAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767994(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_ShapeLinkAddedEventHandler ShapeLinkAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeLinkAddedEvent += value;
			}
			remove
			{
				_ShapeLinkAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_ShapeLinkDeletedEventHandler _ShapeLinkDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765058(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_ShapeLinkDeletedEventHandler ShapeLinkDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ShapeLinkDeletedEvent += value;
			}
			remove
			{
				_ShapeLinkDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event InvisibleApp_AfterRemoveHiddenInformationEventHandler _AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767141(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event InvisibleApp_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent
		{
			add
			{
				CreateEventBridge();
				_AfterRemoveHiddenInformationEvent += value;
			}
			remove
			{
				_AfterRemoveHiddenInformationEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event InvisibleApp_ContainerRelationshipAddedEventHandler _ContainerRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765415(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event InvisibleApp_ContainerRelationshipAddedEventHandler ContainerRelationshipAddedEvent
		{
			add
			{
				CreateEventBridge();
				_ContainerRelationshipAddedEvent += value;
			}
			remove
			{
				_ContainerRelationshipAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event InvisibleApp_ContainerRelationshipDeletedEventHandler _ContainerRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766762(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event InvisibleApp_ContainerRelationshipDeletedEventHandler ContainerRelationshipDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_ContainerRelationshipDeletedEvent += value;
			}
			remove
			{
				_ContainerRelationshipDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event InvisibleApp_CalloutRelationshipAddedEventHandler _CalloutRelationshipAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768554(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event InvisibleApp_CalloutRelationshipAddedEventHandler CalloutRelationshipAddedEvent
		{
			add
			{
				CreateEventBridge();
				_CalloutRelationshipAddedEvent += value;
			}
			remove
			{
				_CalloutRelationshipAddedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event InvisibleApp_CalloutRelationshipDeletedEventHandler _CalloutRelationshipDeletedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765438(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event InvisibleApp_CalloutRelationshipDeletedEventHandler CalloutRelationshipDeletedEvent
		{
			add
			{
				CreateEventBridge();
				_CalloutRelationshipDeletedEvent += value;
			}
			remove
			{
				_CalloutRelationshipDeletedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 14,15,16
		/// </summary>
		private event InvisibleApp_RuleSetValidatedEventHandler _RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766746(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event InvisibleApp_RuleSetValidatedEventHandler RuleSetValidatedEvent
		{
			add
			{
				CreateEventBridge();
				_RuleSetValidatedEvent += value;
			}
			remove
			{
				_RuleSetValidatedEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event InvisibleApp_QueryCancelReplaceShapesEventHandler _QueryCancelReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual event InvisibleApp_QueryCancelReplaceShapesEventHandler QueryCancelReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_QueryCancelReplaceShapesEvent += value;
			}
			remove
			{
				_QueryCancelReplaceShapesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event InvisibleApp_ReplaceShapesCanceledEventHandler _ReplaceShapesCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual event InvisibleApp_ReplaceShapesCanceledEventHandler ReplaceShapesCanceledEvent
		{
			add
			{
				CreateEventBridge();
				_ReplaceShapesCanceledEvent += value;
			}
			remove
			{
				_ReplaceShapesCanceledEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event InvisibleApp_BeforeReplaceShapesEventHandler _BeforeReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual event InvisibleApp_BeforeReplaceShapesEventHandler BeforeReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_BeforeReplaceShapesEvent += value;
			}
			remove
			{
				_BeforeReplaceShapesEvent -= value;
			}
		}

		/// <summary>
		/// SupportByVersion Visio, 15, 16
		/// </summary>
		private event InvisibleApp_AfterReplaceShapesEventHandler _AfterReplaceShapesEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual event InvisibleApp_AfterReplaceShapesEventHandler AfterReplaceShapesEvent
		{
			add
			{
				CreateEventBridge();
				_AfterReplaceShapesEvent += value;
			}
			remove
			{
				_AfterReplaceShapesEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.VisioApi.Behind.EventContracts.EApplication_SinkHelper.Id);


			if (NetOffice.VisioApi.Behind.EventContracts.EApplication_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eApplication_SinkHelper = new NetOffice.VisioApi.Behind.EventContracts.EApplication_SinkHelper(this, _connectPoint);
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
			if( null != _eApplication_SinkHelper)
			{
				_eApplication_SinkHelper.Dispose();
				_eApplication_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

