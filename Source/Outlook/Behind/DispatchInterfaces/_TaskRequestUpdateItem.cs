using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OutlookApi;

namespace NetOffice.OutlookApi.Behind
{
	/// <summary>
	/// DispatchInterface _TaskRequestUpdateItem 
	/// SupportByVersion Outlook, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _TaskRequestUpdateItem : COMObject, NetOffice.OutlookApi._TaskRequestUpdateItem
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
                    _contractType = typeof(NetOffice.OutlookApi._TaskRequestUpdateItem);
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
                    _type = typeof(_TaskRequestUpdateItem);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _TaskRequestUpdateItem() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860369.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Application>(this, "Application");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865635.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlObjectClass>(this, "Class");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861305.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._NameSpace>(this, "Session");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869045.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868920.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Actions Actions
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Actions>(this, "Actions", typeof(NetOffice.OutlookApi.Actions));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867368.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Attachments Attachments
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Attachments>(this, "Attachments", typeof(NetOffice.OutlookApi.Attachments));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866726.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860977.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868336.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867806.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869808.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string ConversationIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationIndex");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866233.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string ConversationTopic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationTopic");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866484.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual DateTime CreationTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "CreationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866184.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string EntryID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "EntryID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869348.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.FormDescription FormDescription
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.FormDescription>(this, "FormDescription", typeof(NetOffice.OutlookApi.FormDescription));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867859.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Inspector GetInspector
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OutlookApi._Inspector>(this, "GetInspector");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867113.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867851.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual DateTime LastModificationTime
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteDateTimePropertyGet(this, "LastModificationTime");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual object MAPIOBJECT
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "MAPIOBJECT");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863006.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869574.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865991.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862687.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual Int32 OutlookInternalVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "OutlookInternalVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861845.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual string OutlookVersion
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OutlookVersion");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868894.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual bool Saved
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Saved");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865033.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868349.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual Int32 Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Size");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867137.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff870156.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866734.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.UserProperties UserProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.UserProperties>(this, "UserProperties", typeof(NetOffice.OutlookApi.UserProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Links Links
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Links>(this, "Links", typeof(NetOffice.OutlookApi.Links));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869304.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlDownloadState DownloadState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlDownloadState>(this, "DownloadState");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866469.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.ItemProperties ItemProperties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.ItemProperties>(this, "ItemProperties", typeof(NetOffice.OutlookApi.ItemProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868928.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Enums.OlRemoteStatus MarkForDownload
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OutlookApi.Enums.OlRemoteStatus>(this, "MarkForDownload");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkForDownload", value);
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868968.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual bool IsConflict
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsConflict");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860631.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual bool AutoResolvedWinner
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoResolvedWinner");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865793.aspx </remarks>
		[SupportByVersion("Outlook", 11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.Conflicts Conflicts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OutlookApi.Conflicts>(this, "Conflicts", typeof(NetOffice.OutlookApi.Conflicts));
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861271.aspx </remarks>
		[SupportByVersion("Outlook", 12,14,15,16)]
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869650.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual string ConversationID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConversationID");
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865341.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		public virtual object RTFBody
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RTFBody");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RTFBody", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867635.aspx </remarks>
		/// <param name="saveMode">NetOffice.OutlookApi.Enums.OlInspectorClose saveMode</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Close(NetOffice.OutlookApi.Enums.OlInspectorClose saveMode)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Close", saveMode);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865812.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual object Copy()
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862499.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Delete()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869400.aspx </remarks>
		/// <param name="modal">optional object modal</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Display(object modal)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display", modal);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869400.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Display()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Display");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868531.aspx </remarks>
		/// <param name="destFldr">NetOffice.OutlookApi.MAPIFolder destFldr</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual object Move(NetOffice.OutlookApi.MAPIFolder destFldr)
		{
			return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Move", destFldr);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff867559.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void PrintOut()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "PrintOut");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863088.aspx </remarks>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void Save()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Save");
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866972.aspx </remarks>
		/// <param name="path">string path</param>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void SaveAs(string path, object type)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", path, type);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff866972.aspx </remarks>
		/// <param name="path">string path</param>
		[CustomMethod]
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual void SaveAs(string path)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SaveAs", path);
		}

		/// <summary>
		/// SupportByVersion Outlook 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff868698.aspx </remarks>
		/// <param name="addToTaskList">bool addToTaskList</param>
		[SupportByVersion("Outlook", 9,10,11,12,14,15,16)]
		public virtual NetOffice.OutlookApi.TaskItem GetAssociatedTask(bool addToTaskList)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OutlookApi.TaskItem>(this, "GetAssociatedTask", typeof(NetOffice.OutlookApi.TaskItem), addToTaskList);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861912.aspx </remarks>
		[SupportByVersion("Outlook", 10,11,12,14,15,16)]
		public virtual void ShowCategoriesDialog()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ShowCategoriesDialog");
		}

		/// <summary>
		/// SupportByVersion Outlook 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff869321.aspx </remarks>
		[SupportByVersion("Outlook", 14,15,16)]
		[BaseResult]
		public virtual NetOffice.OutlookApi._Conversation GetConversation()
		{
			return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OutlookApi._Conversation>(this, "GetConversation");
		}

		#endregion

		#pragma warning restore
	}
}


