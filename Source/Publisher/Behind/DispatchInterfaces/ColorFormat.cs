using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface ColorFormat 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ColorFormat : COMObject, NetOffice.PublisherApi.ColorFormat
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
                    _contractType = typeof(NetOffice.PublisherApi.ColorFormat);
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
                    _type = typeof(ColorFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ColorFormat() : base()
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
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 BaseRGB
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BaseRGB");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BaseRGB", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorCMYK CMYK
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorCMYK>(this, "CMYK", typeof(NetOffice.PublisherApi.ColorCMYK));
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
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 RGB
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "RGB");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RGB", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbSchemeColorIndex SchemeColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbSchemeColorIndex>(this, "SchemeColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SchemeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Ink
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Ink");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Ink", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single TintAndShade
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "TintAndShade");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "TintAndShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.Enums.PbColorType Type
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PublisherApi.Enums.PbColorType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Single Transparency
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Transparency");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Transparency", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorCMYK BaseCMYK
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorCMYK>(this, "BaseCMYK", typeof(NetOffice.PublisherApi.ColorCMYK));
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


