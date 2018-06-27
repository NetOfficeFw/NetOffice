using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.VisioApi.Behind
{
	/// <summary>
	/// CoClass Document 
	/// SupportByVersion Visio, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/ff769268(v=office.14).aspx </remarks>
	[SupportByVersion("Visio", 11,12,14,15,16)]
	[EntityType(EntityType.IsCoClass)]
    [ComEventContract(typeof(NetOffice.VisioApi.EventContracts.EDocument))]
    public class Document : IVDocument, NetOffice.VisioApi.Document
    {
		#pragma warning disable

		#region Fields
		
		private NetRuntimeSystem.Runtime.InteropServices.ComTypes.IConnectionPoint _connectPoint;
		private string _activeSinkId;
        private static Type _type;
        private NetOffice.VisioApi.Behind.EventContracts.EDocument_SinkHelper _eDocument_SinkHelper;
	
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
                    _contractType = typeof(NetOffice.VisioApi.Document);
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
                    _type = typeof(Document);
                return _type;
            }
        }
        
        #endregion
        		
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Document() : base()
		{

		}

		#endregion

		#region Static CoClass Methods
		#endregion

		#region Events

		/// <summary>
		/// SupportByVersion Visio, 11,12,14,15,16
		/// </summary>
		private event Document_DocumentOpenedEventHandler _DocumentOpenedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765849(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentOpenedEventHandler DocumentOpenedEvent
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
		private event Document_DocumentCreatedEventHandler _DocumentCreatedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766513(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentCreatedEventHandler DocumentCreatedEvent
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
		private event Document_DocumentSavedEventHandler _DocumentSavedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766220(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentSavedEventHandler DocumentSavedEvent
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
		private event Document_DocumentSavedAsEventHandler _DocumentSavedAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765909(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentSavedAsEventHandler DocumentSavedAsEvent
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
		private event Document_DocumentChangedEventHandler _DocumentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765974(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentChangedEventHandler DocumentChangedEvent
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
		private event Document_BeforeDocumentCloseEventHandler _BeforeDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768703(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeDocumentCloseEventHandler BeforeDocumentCloseEvent
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
		private event Document_StyleAddedEventHandler _StyleAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768761(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_StyleAddedEventHandler StyleAddedEvent
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
		private event Document_StyleChangedEventHandler _StyleChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765506(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_StyleChangedEventHandler StyleChangedEvent
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
		private event Document_BeforeStyleDeleteEventHandler _BeforeStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768595(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeStyleDeleteEventHandler BeforeStyleDeleteEvent
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
		private event Document_MasterAddedEventHandler _MasterAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766415(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_MasterAddedEventHandler MasterAddedEvent
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
		private event Document_MasterChangedEventHandler _MasterChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766463(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_MasterChangedEventHandler MasterChangedEvent
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
		private event Document_BeforeMasterDeleteEventHandler _BeforeMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766537(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeMasterDeleteEventHandler BeforeMasterDeleteEvent
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
		private event Document_PageAddedEventHandler _PageAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765970(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_PageAddedEventHandler PageAddedEvent
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
		private event Document_PageChangedEventHandler _PageChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767800(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_PageChangedEventHandler PageChangedEvent
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
		private event Document_BeforePageDeleteEventHandler _BeforePageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768594(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforePageDeleteEventHandler BeforePageDeleteEvent
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
		private event Document_ShapeAddedEventHandler _ShapeAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768510(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_ShapeAddedEventHandler ShapeAddedEvent
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
		private event Document_BeforeSelectionDeleteEventHandler _BeforeSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765647(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeSelectionDeleteEventHandler BeforeSelectionDeleteEvent
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
		private event Document_RunModeEnteredEventHandler _RunModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767371(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_RunModeEnteredEventHandler RunModeEnteredEvent
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
		private event Document_DesignModeEnteredEventHandler _DesignModeEnteredEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768278(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DesignModeEnteredEventHandler DesignModeEnteredEvent
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
		private event Document_BeforeDocumentSaveEventHandler _BeforeDocumentSaveEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765091(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeDocumentSaveEventHandler BeforeDocumentSaveEvent
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
		private event Document_BeforeDocumentSaveAsEventHandler _BeforeDocumentSaveAsEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766754(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeDocumentSaveAsEventHandler BeforeDocumentSaveAsEvent
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
		private event Document_QueryCancelDocumentCloseEventHandler _QueryCancelDocumentCloseEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768645(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelDocumentCloseEventHandler QueryCancelDocumentCloseEvent
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
		private event Document_DocumentCloseCanceledEventHandler _DocumentCloseCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769024(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_DocumentCloseCanceledEventHandler DocumentCloseCanceledEvent
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
		private event Document_QueryCancelStyleDeleteEventHandler _QueryCancelStyleDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765142(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelStyleDeleteEventHandler QueryCancelStyleDeleteEvent
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
		private event Document_StyleDeleteCanceledEventHandler _StyleDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768734(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_StyleDeleteCanceledEventHandler StyleDeleteCanceledEvent
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
		private event Document_QueryCancelMasterDeleteEventHandler _QueryCancelMasterDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767941(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelMasterDeleteEventHandler QueryCancelMasterDeleteEvent
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
		private event Document_MasterDeleteCanceledEventHandler _MasterDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768689(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_MasterDeleteCanceledEventHandler MasterDeleteCanceledEvent
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
		private event Document_QueryCancelPageDeleteEventHandler _QueryCancelPageDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768471(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelPageDeleteEventHandler QueryCancelPageDeleteEvent
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
		private event Document_PageDeleteCanceledEventHandler _PageDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff769018(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_PageDeleteCanceledEventHandler PageDeleteCanceledEvent
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
		private event Document_ShapeParentChangedEventHandler _ShapeParentChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765087(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_ShapeParentChangedEventHandler ShapeParentChangedEvent
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
		private event Document_BeforeShapeTextEditEventHandler _BeforeShapeTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768822(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_BeforeShapeTextEditEventHandler BeforeShapeTextEditEvent
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
		private event Document_ShapeExitedTextEditEventHandler _ShapeExitedTextEditEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff767328(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_ShapeExitedTextEditEventHandler ShapeExitedTextEditEvent
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
		private event Document_QueryCancelSelectionDeleteEventHandler _QueryCancelSelectionDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766813(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelSelectionDeleteEventHandler QueryCancelSelectionDeleteEvent
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
		private event Document_SelectionDeleteCanceledEventHandler _SelectionDeleteCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766126(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_SelectionDeleteCanceledEventHandler SelectionDeleteCanceledEvent
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
		private event Document_QueryCancelUngroupEventHandler _QueryCancelUngroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768683(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelUngroupEventHandler QueryCancelUngroupEvent
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
		private event Document_UngroupCanceledEventHandler _UngroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768784(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_UngroupCanceledEventHandler UngroupCanceledEvent
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
		private event Document_QueryCancelConvertToGroupEventHandler _QueryCancelConvertToGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765304(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_QueryCancelConvertToGroupEventHandler QueryCancelConvertToGroupEvent
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
		private event Document_ConvertToGroupCanceledEventHandler _ConvertToGroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 11 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765971(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 11,12,14,15,16)]
		public virtual event Document_ConvertToGroupCanceledEventHandler ConvertToGroupCanceledEvent
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
		/// SupportByVersion Visio, 12,14,15,16
		/// </summary>
		private event Document_QueryCancelGroupEventHandler _QueryCancelGroupEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765279(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_QueryCancelGroupEventHandler QueryCancelGroupEvent
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
		private event Document_GroupCanceledEventHandler _GroupCanceledEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765327(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_GroupCanceledEventHandler GroupCanceledEvent
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
		private event Document_ShapeDataGraphicChangedEventHandler _ShapeDataGraphicChangedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff765116(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_ShapeDataGraphicChangedEventHandler ShapeDataGraphicChangedEvent
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
		private event Document_BeforeDataRecordsetDeleteEventHandler _BeforeDataRecordsetDeleteEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766850(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_BeforeDataRecordsetDeleteEventHandler BeforeDataRecordsetDeleteEvent
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
		private event Document_DataRecordsetAddedEventHandler _DataRecordsetAddedEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766043(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_DataRecordsetAddedEventHandler DataRecordsetAddedEvent
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
		private event Document_AfterRemoveHiddenInformationEventHandler _AfterRemoveHiddenInformationEvent;

		/// <summary>
		/// SupportByVersion Visio 12 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff768457(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 12,14,15,16)]
		public virtual event Document_AfterRemoveHiddenInformationEventHandler AfterRemoveHiddenInformationEvent
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
		private event Document_RuleSetValidatedEventHandler _RuleSetValidatedEvent;

		/// <summary>
		/// SupportByVersion Visio 14 15,16
		/// </summary>
		///<remarks> MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/ff766756(v=office.14).aspx </remarks>
		[SupportByVersion("Visio", 14,15,16)]
		public virtual event Document_RuleSetValidatedEventHandler RuleSetValidatedEvent
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
		private event Document_AfterDocumentMergeEventHandler _AfterDocumentMergeEvent;

		/// <summary>
		/// SupportByVersion Visio 15,16
		/// </summary>
		[SupportByVersion("Visio", 15, 16)]
		public virtual event Document_AfterDocumentMergeEventHandler AfterDocumentMergeEvent
		{
			add
			{
				CreateEventBridge();
				_AfterDocumentMergeEvent += value;
			}
			remove
			{
				_AfterDocumentMergeEvent -= value;
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
				_activeSinkId = SinkHelper.GetConnectionPoint(this, ref _connectPoint, NetOffice.VisioApi.Behind.EventContracts.EDocument_SinkHelper.Id);


			if (NetOffice.VisioApi.Behind.EventContracts.EDocument_SinkHelper.Id.Equals(_activeSinkId, StringComparison.InvariantCultureIgnoreCase))
			{
				_eDocument_SinkHelper = new NetOffice.VisioApi.Behind.EventContracts.EDocument_SinkHelper(this, _connectPoint);
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
			if( null != _eDocument_SinkHelper)
			{
				_eDocument_SinkHelper.Dispose();
				_eDocument_SinkHelper = null;
			}

			_connectPoint = null;
		}
        
        #endregion

		#pragma warning restore
	}
}

