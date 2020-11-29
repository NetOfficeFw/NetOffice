using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _MeetingItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _MeetingItem : COMObject
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
                    _type = typeof(_MeetingItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _MeetingItem(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _MeetingItem(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _MeetingItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Class"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Session"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Actions"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Attachments"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.BillingInformation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Body"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Categories"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Companies"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ConversationIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ConversationTopic"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.CreationTime"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.EntryID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.FormDescription"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.GetInspector"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Importance"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.LastModificationTime"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.MessageClass"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Mileage"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.NoAging"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.OutlookInternalVersion"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.OutlookVersion"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Saved"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Sensitivity"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Size"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Subject"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.UnRead"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.UserProperties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.AutoForwarded"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool AutoForwarded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoForwarded");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoForwarded", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.DeferredDeliveryTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime DeferredDeliveryTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "DeferredDeliveryTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeferredDeliveryTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.DeleteAfterSubmit"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool DeleteAfterSubmit
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DeleteAfterSubmit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DeleteAfterSubmit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ExpiryTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime ExpiryTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ExpiryTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExpiryTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime FlagDueBy
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "FlagDueBy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FlagDueBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FlagRequest
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FlagRequest");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FlagRequest", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlFlagStatus FlagStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlFlagStatus>(this, "FlagStatus");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FlagStatus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.OriginatorDeliveryReportRequested"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool OriginatorDeliveryReportRequested
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OriginatorDeliveryReportRequested");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OriginatorDeliveryReportRequested", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ReceivedTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime ReceivedTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ReceivedTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReceivedTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Recipients"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ReminderSet"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ReminderTime"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime ReminderTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ReminderTime");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReminderTime", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ReplyRecipients"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Recipients ReplyRecipients
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "ReplyRecipients", NetOffice.OutlookApi.Recipients.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SaveSentMessageFolder"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder SaveSentMessageFolder
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi.MAPIFolder>(this, "SaveSentMessageFolder");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "SaveSentMessageFolder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SenderName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string SenderName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Sent"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool Sent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Sent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SentOn"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime SentOn
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "SentOn");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Submitted"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool Submitted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Submitted");
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
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.DownloadState"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ItemProperties"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.MarkForDownload"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.IsConflict"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.MeetingWorkspaceURL"/> </remarks>
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlFlagIcon FlagIcon
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlFlagIcon>(this, "FlagIcon");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FlagIcon", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.AutoResolvedWinner"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Conflicts"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public NetOffice.OutlookApi.Conflicts Conflicts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Conflicts>(this, "Conflicts", NetOffice.OutlookApi.Conflicts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SenderEmailAddress"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public string SenderEmailAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderEmailAddress");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SenderEmailType"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public string SenderEmailType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderEmailType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.PropertyAccessor"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", NetOffice.OutlookApi.PropertyAccessor.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ConversationID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SendUsingAccount"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
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
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.IsLatestVersion"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public bool IsLatestVersion
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsLatestVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.RTFBody"/> </remarks>
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

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.RetentionExpirationDate"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public DateTime RetentionExpirationDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "RetentionExpirationDate");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.RetentionPolicyName"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public string RetentionPolicyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RetentionPolicyName");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Close(method)"/> </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 Factory.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Copy"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Delete"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Display"/> </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display(object modal)
		{
			 Factory.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Display"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Move"/> </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return Factory.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.PrintOut"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Save"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SaveAs"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.SaveAs"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Forward(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MeetingItem Forward()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MeetingItem>(this, "Forward", NetOffice.OutlookApi.MeetingItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.GetAssociatedAppointment"/> </remarks>
		/// <param name="addToCalendar">bool addToCalendar</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.AppointmentItem GetAssociatedAppointment(bool addToCalendar)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.AppointmentItem>(this, "GetAssociatedAppointment", NetOffice.OutlookApi.AppointmentItem.LateBindingApiWrapperType, addToCalendar);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Reply(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MailItem Reply()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "Reply", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ReplyAll(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ReplyAll()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ReplyAll", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.Send(method)"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Send()
		{
			 Factory.ExecuteMethod(this, "Send");
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.ShowCategoriesDialog"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void ShowCategoriesDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.MeetingItem.GetConversation"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Conversation GetConversation()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Conversation>(this, "GetConversation");
		}

		#endregion

		#pragma warning restore
	}
}
