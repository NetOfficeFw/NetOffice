using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _ContactItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _ContactItem : COMObject
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
                    _type = typeof(_ContactItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _ContactItem(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _ContactItem(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _ContactItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Class"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Session"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Actions"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Attachments"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BillingInformation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Body"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Categories"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Companies"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ConversationIndex"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ConversationTopic"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CreationTime"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.EntryID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FormDescription"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.GetInspector"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Importance"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastModificationTime"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MessageClass"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Mileage"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.NoAging"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OutlookInternalVersion"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OutlookVersion"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Saved"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Sensitivity"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Size"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Subject"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.UnRead"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.UserProperties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Account"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Account
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Account");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Account", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Anniversary"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime Anniversary
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "Anniversary");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Anniversary", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.AssistantName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string AssistantName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AssistantName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AssistantName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.AssistantTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string AssistantTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "AssistantTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AssistantTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Birthday"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public DateTime Birthday
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "Birthday");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Birthday", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Business2TelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Business2TelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Business2TelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Business2TelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressCity"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressCity
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressCity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressCity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressCountry"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressCountry
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressCountry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressCountry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressPostalCode"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressPostalCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressPostalCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressPostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressPostOfficeBox"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressPostOfficeBox
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressPostOfficeBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressPostOfficeBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessAddressStreet"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessAddressStreet
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessAddressStreet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessAddressStreet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessFaxNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessFaxNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessFaxNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessFaxNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessHomePage"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessHomePage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessHomePage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessHomePage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string BusinessTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CallbackTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CallbackTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CallbackTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CallbackTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CarTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CarTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CarTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CarTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Children"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Children
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Children");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Children", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CompanyAndFullName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CompanyAndFullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyAndFullName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CompanyLastFirstNoSpace"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CompanyLastFirstNoSpace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyLastFirstNoSpace");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CompanyLastFirstSpaceOnly"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CompanyLastFirstSpaceOnly
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyLastFirstSpaceOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CompanyMainTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CompanyMainTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyMainTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompanyMainTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CompanyName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CompanyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CompanyName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ComputerNetworkName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ComputerNetworkName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ComputerNetworkName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ComputerNetworkName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.CustomerID"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string CustomerID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CustomerID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CustomerID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Department"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Department
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Department");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Department", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email1Address"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email1Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email1Address");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email1Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email1AddressType"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email1AddressType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email1AddressType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email1AddressType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email1DisplayName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email1DisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email1DisplayName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email1DisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email1EntryID"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email1EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email1EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email2Address"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email2Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email2Address");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email2Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email2AddressType"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email2AddressType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email2AddressType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email2AddressType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email2DisplayName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email2DisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email2DisplayName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email2DisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email2EntryID"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email2EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email2EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email3Address"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email3Address
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email3Address");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email3Address", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email3AddressType"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email3AddressType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email3AddressType");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email3AddressType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email3DisplayName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email3DisplayName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email3DisplayName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Email3DisplayName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Email3EntryID"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Email3EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Email3EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FileAs"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FileAs
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FileAs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FileAs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FirstName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FirstName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FirstName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FTPSite"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FTPSite
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FTPSite");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FTPSite", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FullName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FullName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.FullNameAndCompany"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string FullNameAndCompany
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullNameAndCompany");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Gender"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlGender Gender
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlGender>(this, "Gender");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Gender", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.GovernmentIDNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string GovernmentIDNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "GovernmentIDNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GovernmentIDNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Hobby"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Hobby
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Hobby");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hobby", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Home2TelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Home2TelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Home2TelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Home2TelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressCity"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressCity
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressCity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressCity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressCountry"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressCountry
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressCountry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressCountry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressPostalCode"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressPostalCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressPostalCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressPostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressPostOfficeBox"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressPostOfficeBox
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressPostOfficeBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressPostOfficeBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeAddressStreet"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeAddressStreet
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeAddressStreet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeAddressStreet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeFaxNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeFaxNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeFaxNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeFaxNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HomeTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string HomeTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HomeTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HomeTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Initials"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Initials
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Initials");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Initials", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.InternetFreeBusyAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string InternetFreeBusyAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "InternetFreeBusyAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InternetFreeBusyAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ISDNNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ISDNNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ISDNNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ISDNNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.JobTitle"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string JobTitle
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "JobTitle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "JobTitle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Journal"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public bool Journal
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Journal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Journal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Language"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Language
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Language");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Language", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstAndSuffix"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastFirstAndSuffix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstAndSuffix");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstNoSpace"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastFirstNoSpace
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstNoSpace");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstNoSpaceCompany"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastFirstNoSpaceCompany
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstNoSpaceCompany");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstSpaceOnly"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastFirstSpaceOnly
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstSpaceOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstSpaceOnlyCompany"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastFirstSpaceOnlyCompany
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstSpaceOnlyCompany");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LastName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastNameAndFirstName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string LastNameAndFirstName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastNameAndFirstName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressCity"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressCity
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressCity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressCity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressCountry"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressCountry
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressCountry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressCountry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressPostalCode"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressPostalCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressPostalCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressPostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressPostOfficeBox"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressPostOfficeBox
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressPostOfficeBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressPostOfficeBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MailingAddressStreet"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MailingAddressStreet
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MailingAddressStreet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailingAddressStreet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ManagerName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ManagerName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ManagerName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ManagerName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MiddleName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MiddleName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MiddleName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MiddleName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MobileTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string MobileTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MobileTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MobileTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.NetMeetingAlias"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NetMeetingAlias
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NetMeetingAlias");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NetMeetingAlias", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.NetMeetingServer"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.NickName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string NickName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NickName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NickName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OfficeLocation"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OfficeLocation
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OfficeLocation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OfficeLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OrganizationalIDNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OrganizationalIDNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OrganizationalIDNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OrganizationalIDNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressCity"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressCity
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressCity");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressCity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressCountry"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressCountry
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressCountry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressCountry", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressPostalCode"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressPostalCode
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressPostalCode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressPostalCode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressPostOfficeBox"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressPostOfficeBox
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressPostOfficeBox");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressPostOfficeBox", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressState"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressState
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherAddressStreet"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherAddressStreet
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherAddressStreet");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherAddressStreet", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherFaxNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherFaxNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherFaxNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherFaxNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.OtherTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string OtherTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OtherTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OtherTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.PagerNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string PagerNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PagerNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PagerNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.PersonalHomePage"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string PersonalHomePage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PersonalHomePage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PersonalHomePage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.PrimaryTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string PrimaryTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrimaryTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrimaryTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Profession"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Profession
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Profession");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Profession", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.RadioTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string RadioTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RadioTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RadioTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReferredBy"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string ReferredBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReferredBy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReferredBy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.SelectedMailingAddress"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlMailingAddress SelectedMailingAddress
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlMailingAddress>(this, "SelectedMailingAddress");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SelectedMailingAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Spouse"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Spouse
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Spouse");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Spouse", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Suffix"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Suffix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Suffix");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Suffix", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TelexNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string TelexNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TelexNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TelexNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Title"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string Title
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Title");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Title", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TTYTDDTelephoneNumber"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string TTYTDDTelephoneNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TTYTDDTelephoneNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TTYTDDTelephoneNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.User1"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string User1
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "User1");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "User1", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.User2"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string User2
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "User2");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "User2", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.User3"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string User3
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "User3");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "User3", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.User4"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string User4
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "User4");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "User4", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string UserCertificate
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserCertificate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserCertificate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.WebPage"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string WebPage
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WebPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WebPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.YomiCompanyName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string YomiCompanyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiCompanyName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiCompanyName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.YomiFirstName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string YomiFirstName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiFirstName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiFirstName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.YomiLastName"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public string YomiLastName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "YomiLastName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "YomiLastName", value);
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ItemProperties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.LastFirstNoSpaceAndSuffix"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public string LastFirstNoSpaceAndSuffix
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LastFirstNoSpaceAndSuffix");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.DownloadState"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.IMAddress"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public string IMAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "IMAddress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IMAddress", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MarkForDownload"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.IsConflict"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.AutoResolvedWinner"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Conflicts"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.HasPicture"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public bool HasPicture
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasPicture");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.PropertyAccessor"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TaskSubject"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string TaskSubject
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TaskSubject");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskSubject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TaskDueDate"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime TaskDueDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "TaskDueDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskDueDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TaskStartDate"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime TaskStartDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "TaskStartDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskStartDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.TaskCompletedDate"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime TaskCompletedDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "TaskCompletedDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TaskCompletedDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ToDoTaskOrdinal"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime ToDoTaskOrdinal
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ToDoTaskOrdinal");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ToDoTaskOrdinal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReminderOverrideDefault"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReminderPlaySound"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReminderSet"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReminderSoundFile"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ReminderTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.IsMarkedAsTask"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IsMarkedAsTask
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsMarkedAsTask");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessCardLayoutXml"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string BusinessCardLayoutXml
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BusinessCardLayoutXml");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BusinessCardLayoutXml", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.BusinessCardType"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlBusinessCardType BusinessCardType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlBusinessCardType>(this, "BusinessCardType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ConversationID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.RTFBody"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Close(method)"/> </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 Factory.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Copy"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Delete"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Display"/> </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display(object modal)
		{
			 Factory.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Display"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Move"/> </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return Factory.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.PrintOut"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.Save"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.SaveAs"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.SaveAs"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ForwardAsVcard"/> </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsVcard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsVcard", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ShowCategoriesDialog"/> </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void ShowCategoriesDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.AddPicture"/> </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void AddPicture(string path)
		{
			 Factory.ExecuteMethod(this, "AddPicture", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.RemovePicture"/> </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void RemovePicture()
		{
			 Factory.ExecuteMethod(this, "RemovePicture");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ForwardAsBusinessCard"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsBusinessCard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsBusinessCard", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ShowBusinessCardEditor"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ShowBusinessCardEditor()
		{
			 Factory.ExecuteMethod(this, "ShowBusinessCardEditor");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.SaveBusinessCardImage"/> </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SaveBusinessCardImage(string path)
		{
			 Factory.ExecuteMethod(this, "SaveBusinessCardImage", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ShowCheckPhoneDialog"/> </remarks>
		/// <param name="phoneNumber">NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ShowCheckPhoneDialog(NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber)
		{
			 Factory.ExecuteMethod(this, "ShowCheckPhoneDialog", phoneNumber);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.MarkAsTask"/> </remarks>
		/// <param name="markInterval">NetOffice.OutlookApi.Enums.OlMarkInterval markInterval</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void MarkAsTask(NetOffice.OutlookApi.Enums.OlMarkInterval markInterval)
		{
			 Factory.ExecuteMethod(this, "MarkAsTask", markInterval);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ClearTaskFlag"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ClearTaskFlag()
		{
			 Factory.ExecuteMethod(this, "ClearTaskFlag");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.ResetBusinessCard"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ResetBusinessCard()
		{
			 Factory.ExecuteMethod(this, "ResetBusinessCard");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.AddBusinessCardLogoPicture"/> </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void AddBusinessCardLogoPicture(string path)
		{
			 Factory.ExecuteMethod(this, "AddBusinessCardLogoPicture", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.ContactItem.GetConversation"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Conversation GetConversation()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Conversation>(this, "GetConversation");
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.contactitem.showcheckfullnamedialog"/> </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		public void ShowCheckFullNameDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCheckFullNameDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.contactitem.showcheckaddressdialog"/> </remarks>
		/// <param name="mailingAddress">NetOffice.OutlookApi.Enums.OlMailingAddress mailingAddress</param>
		[SupportByVersion("Outlook", 15, 16)]
		public void ShowCheckAddressDialog(NetOffice.OutlookApi.Enums.OlMailingAddress mailingAddress)
		{
			 Factory.ExecuteMethod(this, "ShowCheckAddressDialog", mailingAddress);
		}

		#endregion

		#pragma warning restore
	}
}
