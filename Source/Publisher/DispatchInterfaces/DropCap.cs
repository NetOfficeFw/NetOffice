using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PublisherApi
{
	/// <summary>
	/// DispatchInterface DropCap 
	/// SupportByVersion Publisher, 14,15,16
	/// </summary>
	[SupportByVersion("Publisher", 14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class DropCap : COMObject
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
                    _type = typeof(DropCap);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public DropCap(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public DropCap(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public DropCap(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.Application>(this, "Application", NetOffice.PublisherApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.PublisherApi.ColorFormat FontColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.PublisherApi.ColorFormat>(this, "FontColor", NetOffice.PublisherApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState FontBold
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FontBold");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState FontItalic
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 LinesUp
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LinesUp");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LinesUp", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Size
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Size");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public Int32 Span
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Span");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Span", value);
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
		public void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold, object italic)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", new object[]{ linesUp, size, span, fontName, bold, italic });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyCustomDropCap()
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap");
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyCustomDropCap(object linesUp)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", linesUp);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyCustomDropCap(object linesUp, object size)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size);
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		/// <param name="linesUp">optional Int32 LinesUp = 0</param>
		/// <param name="size">optional Int32 Size = 5</param>
		/// <param name="span">optional Int32 Span = 1</param>
		[CustomMethod]
		[SupportByVersion("Publisher", 14,15,16)]
		public void ApplyCustomDropCap(object linesUp, object size, object span)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size, span);
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
		public void ApplyCustomDropCap(object linesUp, object size, object span, object fontName)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", linesUp, size, span, fontName);
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
		public void ApplyCustomDropCap(object linesUp, object size, object span, object fontName, object bold)
		{
			 Factory.ExecuteMethod(this, "ApplyCustomDropCap", new object[]{ linesUp, size, span, fontName, bold });
		}

		/// <summary>
		/// SupportByVersion Publisher 14, 15, 16
		/// </summary>
		[SupportByVersion("Publisher", 14,15,16)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		#endregion

		#pragma warning restore
	}
}
