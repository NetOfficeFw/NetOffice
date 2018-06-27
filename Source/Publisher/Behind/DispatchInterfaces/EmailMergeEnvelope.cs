using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface EmailMergeEnvelope 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class EmailMergeEnvelope : COMObject, NetOffice.PublisherApi.EmailMergeEnvelope
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
                    _contractType = typeof(NetOffice.PublisherApi.EmailMergeEnvelope);
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
                    _type = typeof(EmailMergeEnvelope);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public EmailMergeEnvelope() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", typeof(NetOffice.PublisherApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeDataField To
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataField>(this, "To", typeof(NetOffice.PublisherApi.MailMergeDataField));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "To", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.MailMergeDataField Cc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.MailMergeDataField>(this, "Cc", typeof(NetOffice.PublisherApi.MailMergeDataField));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Cc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
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
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual string Bcc
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Bcc");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Bcc", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Attachments Attachemts
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Attachments>(this, "Attachemts", typeof(NetOffice.PublisherApi.Attachments));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.pbEmailMergePriority Priority
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.pbEmailMergePriority>(this, "Priority");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Priority", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


