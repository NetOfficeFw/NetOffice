using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface TextFrame2 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.ExcelApi.TextFrame2")]
    public class TextFrame2 : _IMsoDispObj
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
                    _type = typeof(TextFrame2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public TextFrame2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public TextFrame2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public TextFrame2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.Parent"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.MarginBottom"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single MarginBottom
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginBottom");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginBottom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.MarginLeft"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single MarginLeft
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginLeft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginLeft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.MarginRight"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single MarginRight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginRight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginRight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.MarginTop"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single MarginTop
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "MarginTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarginTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.Orientation"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.HorizontalAnchor"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoHorizontalAnchor HorizontalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoHorizontalAnchor>(this, "HorizontalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.VerticalAnchor"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoVerticalAnchor VerticalAnchor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoVerticalAnchor>(this, "VerticalAnchor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalAnchor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.PathFormat"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPathFormat PathFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPathFormat>(this, "PathFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PathFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.WarpFormat"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoWarpFormat WarpFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoWarpFormat>(this, "WarpFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WarpFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.WordArtformat"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoPresetTextEffect WordArtformat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPresetTextEffect>(this, "WordArtformat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WordArtformat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.WordWrap"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState WordWrap
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "WordWrap");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "WordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.AutoSize"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoAutoSize AutoSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoAutoSize>(this, "AutoSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutoSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.ThreeD"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ThreeDFormat ThreeD
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ThreeDFormat>(this, "ThreeD", NetOffice.OfficeApi.ThreeDFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.HasText"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState HasText
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "HasText");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.TextRange"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextRange2 TextRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextRange2>(this, "TextRange", NetOffice.OfficeApi.TextRange2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.Column"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.TextColumn2 Column
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.TextColumn2>(this, "Column", NetOffice.OfficeApi.TextColumn2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.Ruler"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Ruler2 Ruler
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Ruler2>(this, "Ruler", NetOffice.OfficeApi.Ruler2.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.NoTextRotation"/> </remarks>
		[SupportByVersion("Office", 14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState NoTextRotation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "NoTextRotation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NoTextRotation", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.TextFrame2.DeleteText"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public void DeleteText()
		{
			 Factory.ExecuteMethod(this, "DeleteText");
		}

		#endregion

		#pragma warning restore
	}
}
