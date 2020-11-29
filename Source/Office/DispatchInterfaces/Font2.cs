using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Font2 
	/// SupportByVersion Office, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2"/> </remarks>
	[SupportByVersion("Office", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Font2 : _IMsoDispObj
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
                    _type = typeof(Font2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Font2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Font2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Font2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Bold"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Bold
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Bold");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Italic"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Italic
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Italic");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Strike"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextStrike Strike
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextStrike>(this, "Strike");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Strike", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Caps"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextCaps Caps
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextCaps>(this, "Caps");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Caps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.AutorotateNumbers"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState AutorotateNumbers
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "AutorotateNumbers");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "AutorotateNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.BaselineOffset"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single BaselineOffset
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BaselineOffset");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BaselineOffset", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Kerning"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single Kerning
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Kerning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Kerning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Size"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single Size
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Size");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Spacing"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public Single Spacing
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Spacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Spacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.UnderlineStyle"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTextUnderlineType UnderlineStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTextUnderlineType>(this, "UnderlineStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UnderlineStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Allcaps"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Allcaps
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Allcaps");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Allcaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.DoubleStrikeThrough"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState DoubleStrikeThrough
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "DoubleStrikeThrough");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DoubleStrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Equalize"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Equalize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Equalize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Equalize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Fill"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.FillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.FillFormat>(this, "Fill", NetOffice.OfficeApi.FillFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Glow"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.GlowFormat Glow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GlowFormat>(this, "Glow", NetOffice.OfficeApi.GlowFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Reflection"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ReflectionFormat Reflection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ReflectionFormat>(this, "Reflection", NetOffice.OfficeApi.ReflectionFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Line"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.LineFormat Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.LineFormat>(this, "Line", NetOffice.OfficeApi.LineFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Shadow"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ShadowFormat Shadow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ShadowFormat>(this, "Shadow", NetOffice.OfficeApi.ShadowFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Highlight"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ColorFormat Highlight
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "Highlight", NetOffice.OfficeApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.UnderlineColor"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.ColorFormat UnderlineColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "UnderlineColor", NetOffice.OfficeApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Smallcaps"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Smallcaps
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Smallcaps");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Smallcaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.SoftEdgeFormat"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoSoftEdgeType SoftEdgeFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSoftEdgeType>(this, "SoftEdgeFormat");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SoftEdgeFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.StrikeThrough"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState StrikeThrough
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "StrikeThrough");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Subscript"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Subscript
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Subscript");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Subscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Superscript"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Superscript
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Superscript");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Superscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.WordArtformat"/> </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Embeddable"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Embeddable
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Embeddable");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Embedded"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoTriState Embedded
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Embedded");
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.Name"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.NameAscii"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NameAscii
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameAscii");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameAscii", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.NameComplexScript"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NameComplexScript
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameComplexScript");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameComplexScript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.NameFarEast"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NameFarEast
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameFarEast");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameFarEast", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.Font2.NameOther"/> </remarks>
		[SupportByVersion("Office", 12,14,15,16)]
		public string NameOther
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameOther");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameOther", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
