using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface Mailer 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838452.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Mailer : COMObject, NetOffice.ExcelApi.Mailer
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
                    _type = typeof(Mailer);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Mailer() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194943.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194093.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193540.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822319.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object BCCRecipients
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BCCRecipients");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BCCRecipients", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840725.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object CCRecipients
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CCRecipients");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "CCRecipients", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193252.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object Enclosures
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Enclosures");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Enclosures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834369.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual bool Received
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Received");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835223.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual DateTime SendDateTime
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "SendDateTime");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835870.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual string Sender
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Sender");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835284.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual string Subject
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822901.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object ToRecipients
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ToRecipients");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ToRecipients", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837563.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object WhichAddress
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "WhichAddress");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "WhichAddress", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


