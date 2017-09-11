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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868476.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866948.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868700.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864478.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862161.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865274.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868961.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865615.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868930.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863616.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866064.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868608.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867864.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860429.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865802.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869254.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869965.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861547.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863901.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867548.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869251.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864258.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862406.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869811.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869844.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863365.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868050.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865670.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869955.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860971.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868919.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860640.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860982.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869282.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868769.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867232.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866226.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869143.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860952.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864267.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860975.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862080.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867303.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867885.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866227.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862663.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864290.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869537.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868047.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869476.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867631.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862248.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860654.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863063.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867329.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866006.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860936.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869942.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866426.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867187.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861590.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860739.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863592.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860946.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868579.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868560.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869715.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869918.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866001.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863968.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862471.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863059.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867814.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860335.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869210.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865840.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861802.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869043.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862079.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868311.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862692.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868107.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868798.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868070.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869795.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869394.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869869.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868657.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868030.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866080.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863618.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866054.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868609.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862992.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865046.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868471.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867834.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864221.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866735.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866926.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870077.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860947.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868831.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868668.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868200.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867259.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868867.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860667.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864197.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869794.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867383.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869407.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870105.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863903.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861601.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868472.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869088.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868437.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867645.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868392.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869474.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868199.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862246.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862776.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869119.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868841.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864488.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861307.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860437.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867114.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864728.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869774.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869871.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863910.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867390.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869725.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865797.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870180.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867668.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860704.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862400.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868455.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864219.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864274.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861577.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863071.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869375.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867597.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863421.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869858.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868985.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865639.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869642.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870041.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863344.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870049.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865987.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860673.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860694.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868429.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867607.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868467.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869058.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868859.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860753.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865308.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861318.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870062.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861625.aspx </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 Factory.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861014.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Copy()
		{
			return Factory.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862373.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866773.aspx </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display(object modal)
		{
			 Factory.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866773.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Display()
		{
			 Factory.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869633.aspx </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return Factory.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867897.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void PrintOut()
		{
			 Factory.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862148.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public void Save()
		{
			 Factory.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868220.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868220.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863898.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsVcard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsVcard", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862359.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public void ShowCategoriesDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868443.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void AddPicture(string path)
		{
			 Factory.ExecuteMethod(this, "AddPicture", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868359.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public void RemovePicture()
		{
			 Factory.ExecuteMethod(this, "RemovePicture");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863014.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public NetOffice.OutlookApi.MailItem ForwardAsBusinessCard()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MailItem>(this, "ForwardAsBusinessCard", NetOffice.OutlookApi.MailItem.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867880.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ShowBusinessCardEditor()
		{
			 Factory.ExecuteMethod(this, "ShowBusinessCardEditor");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867386.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void SaveBusinessCardImage(string path)
		{
			 Factory.ExecuteMethod(this, "SaveBusinessCardImage", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863935.aspx </remarks>
		/// <param name="phoneNumber">NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ShowCheckPhoneDialog(NetOffice.OutlookApi.Enums.OlContactPhoneNumber phoneNumber)
		{
			 Factory.ExecuteMethod(this, "ShowCheckPhoneDialog", phoneNumber);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869513.aspx </remarks>
		/// <param name="markInterval">NetOffice.OutlookApi.Enums.OlMarkInterval markInterval</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void MarkAsTask(NetOffice.OutlookApi.Enums.OlMarkInterval markInterval)
		{
			 Factory.ExecuteMethod(this, "MarkAsTask", markInterval);
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861851.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ClearTaskFlag()
		{
			 Factory.ExecuteMethod(this, "ClearTaskFlag");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868370.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void ResetBusinessCard()
		{
			 Factory.ExecuteMethod(this, "ResetBusinessCard");
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866480.aspx </remarks>
		/// <param name="path">string path</param>
		[SupportByVersion("Outlook", 12,14,15,16)]
		public void AddBusinessCardLogoPicture(string path)
		{
			 Factory.ExecuteMethod(this, "AddBusinessCardLogoPicture", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861831.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public NetOffice.OutlookApi._Conversation GetConversation()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Conversation>(this, "GetConversation");
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231065.aspx </remarks>
		[SupportByVersion("Outlook", 15, 16)]
		public void ShowCheckFullNameDialog()
		{
			 Factory.ExecuteMethod(this, "ShowCheckFullNameDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229480.aspx </remarks>
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
