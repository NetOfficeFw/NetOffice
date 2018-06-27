using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PublisherApi;

namespace NetOffice.PublisherApi.Behind
{
	/// <summary>
	/// DispatchInterface DropCap 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DropCap : COMObject, NetOffice.PublisherApi.DropCap
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
                    _contractType = typeof(NetOffice.PublisherApi.DropCap);
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
                    _type = typeof(DropCap);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public DropCap() : base()
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
		public virtual string FontName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.PublisherApi.ColorFormat FontColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "FontColor", typeof(NetOffice.PublisherApi.ColorFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState FontBold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FontBold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual NetOffice.OfficeApi.Enums.MsoTriState FontItalic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FontItalic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 LinesUp
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LinesUp");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LinesUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Size");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual Int32 Span
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Span");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Span", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="bold">optional bool Bold = false</param>
		/// <param name="italic">optional bool Italic = false</param>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold, object italic)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", new object[]{ linesUp, size, span, fontName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", linesUp);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp, object size)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp, object size, object span)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size, span);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp, object size, object span, object fontName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size, span, fontName);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		/// <param name="fontName">optional string FontName = </param>
		/// <param name="bold">optional bool Bold = false</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomDropCap", new object[]{ linesUp, size, span, fontName, bold });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public virtual void Clear()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Clear");
		}

		#endregion

		#pragma warning restore
	}
}


