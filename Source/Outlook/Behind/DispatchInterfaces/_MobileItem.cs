using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _MobileItem 
	/// SupportByVersion Outlook, 14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _MobileItem : COMObject, NetOffice.OutlookApi._MobileItem
	{
		#pragma warning disable

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
                    _contractType = typeof(NetOffice.OutlookApi._MobileItem);
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

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(_MobileItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _MobileItem() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Actions Actions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Actions>(this, "Actions", typeof(NetOffice.OutlookApi.Actions));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Attachments Attachments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Attachments>(this, "Attachments", typeof(NetOffice.OutlookApi.Attachments));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string BillingInformation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BillingInformation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BillingInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string Body
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Body");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Body", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string Categories
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Categories");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Categories", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string Companies
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Companies");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Companies", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ConversationIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ConversationTopic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationTopic");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual DateTime CreationTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "CreationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string EntryID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.FormDescription FormDescription
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.FormDescription>(this, "FormDescription", typeof(NetOffice.OutlookApi.FormDescription));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Inspector GetInspector
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspector>(this, "GetInspector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlImportance Importance
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlImportance>(this, "Importance");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Importance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual DateTime LastModificationTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "LastModificationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object MAPIOBJECT
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string MessageClass
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MessageClass");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MessageClass", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string Mileage
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Mileage");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Mileage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool NoAging
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "NoAging");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoAging", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual Int32 OutlookInternalVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "OutlookInternalVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string OutlookVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlookVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlSensitivity Sensitivity
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlSensitivity>(this, "Sensitivity");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Sensitivity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual Int32 Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string Subject
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Subject");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Subject", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool UnRead
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UnRead");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UnRead", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.UserProperties UserProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserProperties>(this, "UserProperties", typeof(NetOffice.OutlookApi.UserProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string HTMLBody
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "HTMLBody");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HTMLBody", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlMobileFormat MobileFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlMobileFormat>(this, "MobileFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string SMILBody
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SMILBody");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SMILBody", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Recipients Recipients
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "Recipients", typeof(NetOffice.OutlookApi.Recipients));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string To
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "To");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "To", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ReplyRecipientNames
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ReplyRecipientNames");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Recipients ReplyRecipients
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Recipients>(this, "ReplyRecipients", typeof(NetOffice.OutlookApi.Recipients));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool Submitted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Submitted");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.ItemProperties ItemProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ItemProperties>(this, "ItemProperties", typeof(NetOffice.OutlookApi.ItemProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual DateTime ReceivedTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "ReceivedTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.Account SendUsingAccount
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Account>(this, "SendUsingAccount", typeof(NetOffice.OutlookApi.Account));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "SendUsingAccount", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual bool Sent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Sent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual DateTime SentOn
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "SentOn");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.PropertyAccessor PropertyAccessor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.PropertyAccessor>(this, "PropertyAccessor", typeof(NetOffice.OutlookApi.PropertyAccessor));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ReceivedByEntryID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ReceivedByEntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ReceivedByName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ReceivedByName");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string SenderEmailAddress
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SenderEmailAddress");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string SenderEmailType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SenderEmailType");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string SenderName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "SenderName");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual object Copy()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Display(object modal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Display()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SaveAs(string path, object type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", path, type);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void SaveAs(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.MobileItem Reply()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MobileItem>(this, "Reply", typeof(NetOffice.OutlookApi.MobileItem));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.MobileItem ReplyAll()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MobileItem>(this, "ReplyAll", typeof(NetOffice.OutlookApi.MobileItem));
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <param name="forceSend">bool forceSend</param>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual void Send(bool forceSend)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Send", forceSend);
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual NetOffice.OutlookApi.MobileItem Forward()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.MobileItem>(this, "Forward", typeof(NetOffice.OutlookApi.MobileItem));
		}

		#endregion

		#pragma warning restore
	}
}


