using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _AppointmentItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _AppointmentItem : COMObject
	{
		#pragma warning disable

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

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_AppointmentItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _AppointmentItem(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _AppointmentItem(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _AppointmentItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Application"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Class"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Session"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Parent"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Actions"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Actions Actions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Actions>(this, "Actions", NetOffice.OutlookApi.Actions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Attachments"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Attachments Attachments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Attachments>(this, "Attachments", NetOffice.OutlookApi.Attachments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.BillingInformation"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BillingInformation
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BillingInformation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BillingInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Body"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Body
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Body");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Body", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Categories"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Categories
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Categories");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Categories", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Companies"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Companies
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Companies");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Companies", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ConversationIndex"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ConversationIndex
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ConversationTopic"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ConversationTopic
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationTopic");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.CreationTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime CreationTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "CreationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.EntryID"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.FormDescription"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.FormDescription FormDescription
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.FormDescription>(this, "FormDescription", NetOffice.OutlookApi.FormDescription.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.GetInspector"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Inspector GetInspector
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspector>(this, "GetInspector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Importance"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlImportance Importance
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlImportance>(this, "Importance");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Importance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.LastModificationTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime LastModificationTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "LastModificationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MAPIOBJECT
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.MessageClass"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MessageClass
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MessageClass");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MessageClass", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Mileage"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Mileage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Mileage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Mileage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.NoAging"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool NoAging
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NoAging");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoAging", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.OutlookInternalVersion"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 OutlookInternalVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OutlookInternalVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.OutlookVersion"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OutlookVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlookVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Saved"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Sensitivity"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlSensitivity Sensitivity
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlSensitivity>(this, "Sensitivity");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Sensitivity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Size"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Size
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Subject"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Subject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Subject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.UnRead"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool UnRead
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UnRead");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UnRead", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.UserProperties"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.UserProperties UserProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserProperties>(this, "UserProperties", NetOffice.OutlookApi.UserProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.AllDayEvent"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool AllDayEvent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllDayEvent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllDayEvent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.BusyStatus"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlBusyStatus BusyStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlBusyStatus>(this, "BusyStatus");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BusyStatus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Duration"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 Duration
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Duration");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Duration", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.End"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime End
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "End");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "End", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool IsOnlineMeeting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsOnlineMeeting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsOnlineMeeting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.IsRecurring"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool IsRecurring
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsRecurring");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Location"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Location
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Location");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Location", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.MeetingStatus"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlMeetingStatus MeetingStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlMeetingStatus>(this, "MeetingStatus");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MeetingStatus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool NetMeetingAutoStart
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NetMeetingAutoStart");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetMeetingAutoStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NetMeetingOrganizerAlias
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NetMeetingOrganizerAlias");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetMeetingOrganizerAlias", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NetMeetingServer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NetMeetingServer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetMeetingServer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlNetMeetingType NetMeetingType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlNetMeetingType>(this, "NetMeetingType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NetMeetingType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.OptionalAttendees"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OptionalAttendees
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OptionalAttendees");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OptionalAttendees", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Organizer"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Organizer
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Organizer");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Recipients"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipients Recipients
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "Recipients", NetOffice.OutlookApi.Recipients.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.RecurrenceState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlRecurrenceState RecurrenceState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRecurrenceState>(this, "RecurrenceState");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReminderMinutesBeforeStart"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public Int32 ReminderMinutesBeforeStart
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ReminderMinutesBeforeStart");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderMinutesBeforeStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReminderOverrideDefault"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool ReminderOverrideDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReminderOverrideDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderOverrideDefault", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReminderPlaySound"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool ReminderPlaySound
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReminderPlaySound");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderPlaySound", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReminderSet"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool ReminderSet
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReminderSet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderSet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReminderSoundFile"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ReminderSoundFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReminderSoundFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderSoundFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ReplyTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime ReplyTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ReplyTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReplyTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.RequiredAttendees"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string RequiredAttendees
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RequiredAttendees");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RequiredAttendees", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Resources"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Resources
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Resources");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Resources", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ResponseRequested"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool ResponseRequested
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ResponseRequested");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ResponseRequested", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ResponseStatus"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlResponseStatus ResponseStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlResponseStatus>(this, "ResponseStatus");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Start"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime Start
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "Start");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NetMeetingDocPathName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NetMeetingDocPathName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetMeetingDocPathName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NetShowURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NetShowURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetShowURL", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Links Links
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Links>(this, "Links", NetOffice.OutlookApi.Links.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool ConferenceServerAllowExternal
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConferenceServerAllowExternal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConferenceServerAllowExternal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ConferenceServerPassword
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConferenceServerPassword");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConferenceServerPassword", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ItemProperties"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.ItemProperties ItemProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ItemProperties>(this, "ItemProperties", NetOffice.OutlookApi.ItemProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.DownloadState"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDownloadState DownloadState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDownloadState>(this, "DownloadState");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.InternetCodepage"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public Int32 InternetCodepage
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "InternetCodepage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InternetCodepage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.MarkForDownload"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlRemoteStatus MarkForDownload
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRemoteStatus>(this, "MarkForDownload");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkForDownload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.IsConflict"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public bool IsConflict
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsConflict");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.MeetingWorkspaceURL"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public string MeetingWorkspaceURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MeetingWorkspaceURL");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.AutoResolvedWinner"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public bool AutoResolvedWinner
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoResolvedWinner");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Conflicts"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflicts Conflicts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Conflicts>(this, "Conflicts", NetOffice.OutlookApi.Conflicts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.PropertyAccessor"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", NetOffice.OutlookApi.PropertyAccessor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.SendUsingAccount"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Account SendUsingAccount
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Account>(this, "SendUsingAccount", NetOffice.OutlookApi.Account.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SendUsingAccount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.GlobalAppointmentID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string GlobalAppointmentID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "GlobalAppointmentID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ForceUpdateToAllAttendees"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ForceUpdateToAllAttendees
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ForceUpdateToAllAttendees");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForceUpdateToAllAttendees", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.StartUTC"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime StartUTC
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "StartUTC");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartUTC", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.EndUTC"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime EndUTC
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "EndUTC");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EndUTC", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.StartInStartTimeZone"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime StartInStartTimeZone
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "StartInStartTimeZone");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartInStartTimeZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.EndInEndTimeZone"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime EndInEndTimeZone
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "EndInEndTimeZone");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EndInEndTimeZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.StartTimeZone"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._TimeZone StartTimeZone
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._TimeZone>(this, "StartTimeZone");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "StartTimeZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.EndTimeZone"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._TimeZone EndTimeZone
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._TimeZone>(this, "EndTimeZone");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "EndTimeZone", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ConversationID"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public string ConversationID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.RTFBody"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public object RTFBody
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "RTFBody");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "RTFBody", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Close(method)"/> </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 Factory.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Copy"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Delete"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Display"/> </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display(object modal)
		{
			 Factory.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Display"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Move"/> </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return Factory.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.PrintOut"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Save"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.SaveAs"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void SaveAs(string path, object type)
		{
			 Factory.ExecuteMethod(this, "SaveAs", path, type);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.SaveAs"/> </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void SaveAs(string path)
		{
			 Factory.ExecuteMethod(this, "SaveAs", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ClearRecurrencePattern"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void ClearRecurrencePattern()
		{
			 Factory.ExecuteMethod(this, "ClearRecurrencePattern");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ForwardAsVcal"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsVcal()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsVcal", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.GetRecurrencePattern"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.RecurrencePattern GetRecurrencePattern()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.RecurrencePattern>(this, "GetRecurrencePattern", NetOffice.OutlookApi.RecurrencePattern.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Respond"/> </remarks>
		/// <param name="response">NetOffice.OutlookApi.Enums.OlMeetingResponse response</param>
		/// <param name="fNoUI">optional object fNoUI</param>
		/// <param name="fAdditionalTextDialog">optional object fAdditionalTextDialog</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MeetingItem Respond(NetOffice.OutlookApi.Enums.OlMeetingResponse response, object fNoUI, object fAdditionalTextDialog)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MeetingItem>(this, "Respond", NetOffice.OutlookApi.MeetingItem.LateBindingApiWrapperType, response, fNoUI, fAdditionalTextDialog);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Respond"/> </remarks>
		/// <param name="response">NetOffice.OutlookApi.Enums.OlMeetingResponse response</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MeetingItem Respond(NetOffice.OutlookApi.Enums.OlMeetingResponse response)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MeetingItem>(this, "Respond", NetOffice.OutlookApi.MeetingItem.LateBindingApiWrapperType, response);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Respond"/> </remarks>
		/// <param name="response">NetOffice.OutlookApi.Enums.OlMeetingResponse response</param>
		/// <param name="fNoUI">optional object fNoUI</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MeetingItem Respond(NetOffice.OutlookApi.Enums.OlMeetingResponse response, object fNoUI)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MeetingItem>(this, "Respond", NetOffice.OutlookApi.MeetingItem.LateBindingApiWrapperType, response, fNoUI);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.Send(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Send()
		{
			 Factory.ExecuteMethod(this, "Send");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.ShowCategoriesDialog"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void ShowCategoriesDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.GetConversation"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Conversation GetConversation()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Conversation>(this, "GetConversation");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.CopyTo"/> </remarks>
		/// <param name="destinationFolder">NetOffice.OutlookApi.MAPIFolder destinationFolder</param>
		/// <param name="copyOptions">NetOffice.OutlookApi.Enums.OlAppointmentCopyOptions copyOptions</param>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._AppointmentItem CopyTo(NetOffice.OutlookApi.MAPIFolder destinationFolder, NetOffice.OutlookApi.Enums.OlAppointmentCopyOptions copyOptions)
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._AppointmentItem>(this, "CopyTo", destinationFolder, copyOptions);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.AppointmentItem.GetOrganizer"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public NetOffice.OutlookApi.AddressEntry GetOrganizer()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AddressEntry>(this, "GetOrganizer", NetOffice.OutlookApi.AddressEntry.LateBindingApiWrapperType);
		}

		#endregion

		#pragma warning restore
	}
}
