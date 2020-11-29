using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OutlookApi
{
	/// <summary>
	/// DispatchInterface _SharingItem 
	/// SupportByVersion Outlook, 12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _SharingItem : COMObject
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
                    _type = typeof(_SharingItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _SharingItem(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _SharingItem(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _SharingItem(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Application"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Class"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Session"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Parent"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Actions"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Actions Actions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Actions>(this, "Actions", NetOffice.OutlookApi.Actions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Attachments"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Attachments Attachments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Attachments>(this, "Attachments", NetOffice.OutlookApi.Attachments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.BillingInformation"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Body"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Categories"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Companies"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ConversationIndex"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ConversationIndex
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ConversationTopic"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ConversationTopic
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConversationTopic");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.CreationTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime CreationTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "CreationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.EntryID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string EntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.FormDescription"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.FormDescription FormDescription
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.FormDescription>(this, "FormDescription", NetOffice.OutlookApi.FormDescription.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.GetInspector"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Inspector GetInspector
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspector>(this, "GetInspector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Importance"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.LastModificationTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime LastModificationTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "LastModificationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object MAPIOBJECT
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.MessageClass"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Mileage"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.NoAging"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.OutlookInternalVersion"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public Int32 OutlookInternalVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OutlookInternalVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.OutlookVersion"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string OutlookVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OutlookVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Saved"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool Saved
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Sensitivity"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Size"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public Int32 Size
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Subject"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.UnRead"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.UserProperties"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.UserProperties UserProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserProperties>(this, "UserProperties", NetOffice.OutlookApi.UserProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.PropertyAccessor"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RemoteName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string RemoteName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RemoteName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RemoteID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string RemoteID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RemoteID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RemotePath"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string RemotePath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RemotePath");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SharingProviderGuid"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string SharingProviderGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SharingProviderGuid");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SharingProvider"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlSharingProvider SharingProvider
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlSharingProvider>(this, "SharingProvider");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.AllowWriteAccess"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool AllowWriteAccess
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowWriteAccess");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllowWriteAccess", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Type"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlSharingMsgType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlSharingMsgType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RequestedFolder"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDefaultFolders RequestedFolder
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDefaultFolders>(this, "RequestedFolder");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SendUsingAccount"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.AlternateRecipientAllowed"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool AlternateRecipientAllowed
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlternateRecipientAllowed");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlternateRecipientAllowed", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.AutoForwarded"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.BCC"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string BCC
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BCC");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BCC", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.CC"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string CC
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CC");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CC", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.DeferredDeliveryTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.DeleteAfterSubmit"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ExpiryTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.FlagRequest"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.HTMLBody"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string HTMLBody
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "HTMLBody");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HTMLBody", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.OriginatorDeliveryReportRequested"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReadReceiptRequested"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool ReadReceiptRequested
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadReceiptRequested");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadReceiptRequested", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReceivedByEntryID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ReceivedByEntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReceivedByEntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReceivedByName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ReceivedByName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReceivedByName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReceivedOnBehalfOfEntryID"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ReceivedOnBehalfOfEntryID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReceivedOnBehalfOfEntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReceivedOnBehalfOfName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ReceivedOnBehalfOfName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReceivedOnBehalfOfName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReceivedTime"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime ReceivedTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "ReceivedTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RecipientReassignmentProhibited"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool RecipientReassignmentProhibited
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RecipientReassignmentProhibited");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RecipientReassignmentProhibited", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Recipients"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Recipients Recipients
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "Recipients", NetOffice.OutlookApi.Recipients.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReminderOverrideDefault"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReminderPlaySound"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReminderSet"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReminderSoundFile"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReminderTime"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RemoteStatus"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlRemoteStatus RemoteStatus
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRemoteStatus>(this, "RemoteStatus");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RemoteStatus", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReplyRecipientNames"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string ReplyRecipientNames
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ReplyRecipientNames");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReplyRecipients"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Recipients ReplyRecipients
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "ReplyRecipients", NetOffice.OutlookApi.Recipients.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SaveSentMessageFolder"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SenderName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string SenderName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Sent"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool Sent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Sent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SentOn"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public DateTime SentOn
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "SentOn");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SentOnBehalfOfName"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string SentOnBehalfOfName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SentOnBehalfOfName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SentOnBehalfOfName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Submitted"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool Submitted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Submitted");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.To"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string To
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "To");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "To", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ItemProperties"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.ItemProperties ItemProperties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ItemProperties>(this, "ItemProperties", NetOffice.OutlookApi.ItemProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.BodyFormat"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlBodyFormat BodyFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlBodyFormat>(this, "BodyFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BodyFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.DownloadState"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlDownloadState DownloadState
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDownloadState>(this, "DownloadState");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.InternetCodepage"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.MarkForDownload"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.IsConflict"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IsConflict
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsConflict");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.TaskSubject"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.TaskDueDate"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.TaskStartDate"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.TaskCompletedDate"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ToDoTaskOrdinal"/> </remarks>
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
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Conflicts"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SenderEmailAddress"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string SenderEmailAddress
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderEmailAddress");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SenderEmailType"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public string SenderEmailType
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SenderEmailType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool EnableSharedAttachments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EnableSharedAttachments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnableSharedAttachments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Permission"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlPermission Permission
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlPermission>(this, "Permission");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Permission", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.PermissionService"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.Enums.OlPermissionService PermissionService
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlPermissionService>(this, "PermissionService");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PermissionService", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.IsMarkedAsTask"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public bool IsMarkedAsTask
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsMarkedAsTask");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ConversationID"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.PermissionTemplateGuid"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public string PermissionTemplateGuid
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PermissionTemplateGuid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PermissionTemplateGuid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RTFBody"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RetentionPolicyName"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public string RetentionPolicyName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RetentionPolicyName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.RetentionExpirationDate"/> </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public DateTime RetentionExpirationDate
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "RetentionExpirationDate");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Close(method)"/> </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 Factory.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Copy"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Delete"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Display"/> </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Display(object modal)
		{
			 Factory.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Display"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Move"/> </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return Factory.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.PrintOut"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Save"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SaveAs"/> </remarks>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SaveAs(string path, object type)
		{
			 Factory.ExecuteMethod(this, "SaveAs", path, type);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.SaveAs"/> </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SaveAs(string path)
		{
			 Factory.ExecuteMethod(this, "SaveAs", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Allow"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Allow()
		{
			 Factory.ExecuteMethod(this, "Allow");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Deny"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem Deny()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SharingItem>(this, "Deny", NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.OpenSharedFolder"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi.MAPIFolder OpenSharedFolder()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi.MAPIFolder>(this, "OpenSharedFolder");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ClearConversationIndex"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ClearConversationIndex()
		{
			 Factory.ExecuteMethod(this, "ClearConversationIndex");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Forward(method)"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.SharingItem Forward()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.SharingItem>(this, "Forward", NetOffice.OutlookApi.SharingItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Reply(method)"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MailItem Reply()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "Reply", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ReplyAll(method)"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ReplyAll()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ReplyAll", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.Send(method)"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void Send()
		{
			 Factory.ExecuteMethod(this, "Send");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ShowCategoriesDialog"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ShowCategoriesDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.AddBusinessCard"/> </remarks>
		/// <param name="contact">NetOffice.OutlookApi.ContactItem contact</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void AddBusinessCard(NetOffice.OutlookApi.ContactItem contact)
		{
			 Factory.ExecuteMethod(this, "AddBusinessCard", contact);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.MarkAsTask"/> </remarks>
		/// <param name="markInterval">NetOffice.OutlookApi.Enums.OlMarkInterval markInterval</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void MarkAsTask(NetOffice.OutlookApi.Enums.OlMarkInterval markInterval)
		{
			 Factory.ExecuteMethod(this, "MarkAsTask", markInterval);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.ClearTaskFlag"/> </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ClearTaskFlag()
		{
			 Factory.ExecuteMethod(this, "ClearTaskFlag");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Outlook.SharingItem.GetConversation"/> </remarks>
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
