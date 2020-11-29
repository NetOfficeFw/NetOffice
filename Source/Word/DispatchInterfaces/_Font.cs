using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface _Font 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Font : COMObject
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
                    _type = typeof(_Font);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _Font(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _Font(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Font(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Application"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", NetOffice.WordApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Creator"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Creator
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Parent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Duplicate"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Font Duplicate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Font>(this, "Duplicate", NetOffice.WordApi.Font.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Bold"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Bold
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Bold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Italic"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Italic
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Italic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Hidden"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Hidden
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hidden");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.SmallCaps"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SmallCaps
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SmallCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SmallCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.AllCaps"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 AllCaps
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AllCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AllCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.StrikeThrough"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StrikeThrough
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StrikeThrough");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.DoubleStrikeThrough"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 DoubleStrikeThrough
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DoubleStrikeThrough");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DoubleStrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.ColorIndex"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex ColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "ColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Subscript"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Subscript
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Subscript");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Subscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Superscript"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Superscript
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Superscript");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Superscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Underline"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdUnderline Underline
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdUnderline>(this, "Underline");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Size"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Name"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Position"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Position
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Position");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Spacing"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Scaling"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Scaling
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Scaling");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Scaling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Shadow"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Shadow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Shadow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Outline"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Outline
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Outline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Outline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Emboss"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Emboss
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Emboss");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Emboss", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Kerning"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Engrave"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Engrave
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Engrave");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Engrave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdAnimation Animation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAnimation>(this, "Animation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Animation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Borders"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Borders Borders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", NetOffice.WordApi.Borders.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Shading"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Shading Shading
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shading>(this, "Shading", NetOffice.WordApi.Shading.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.EmphasisMark"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEmphasisMark>(this, "EmphasisMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EmphasisMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.DisableCharacterSpaceGrid"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisableCharacterSpaceGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableCharacterSpaceGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableCharacterSpaceGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NameFarEast"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NameAscii"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NameOther"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColor Color
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "Color");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.BoldBi"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BoldBi
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BoldBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BoldBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.ItalicBi"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ItalicBi
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ItalicBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ItalicBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.SizeBi"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single SizeBi
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SizeBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SizeBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NameBi"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string NameBi
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "NameBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NameBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.ColorIndexBi"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex ColorIndexBi
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "ColorIndexBi");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ColorIndexBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.DiacriticColor"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColor DiacriticColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "DiacriticColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DiacriticColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.UnderlineColor"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColor UnderlineColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "UnderlineColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "UnderlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Glow"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.GlowFormat Glow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.GlowFormat>(this, "Glow", NetOffice.WordApi.GlowFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Glow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Reflection"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ReflectionFormat Reflection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReflectionFormat>(this, "Reflection", NetOffice.WordApi.ReflectionFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Reflection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.TextShadow"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ShadowFormat TextShadow
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShadowFormat>(this, "TextShadow", NetOffice.WordApi.ShadowFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "TextShadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Fill"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.FillFormat Fill
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FillFormat>(this, "Fill", NetOffice.WordApi.FillFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Fill", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Line"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.LineFormat Line
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LineFormat>(this, "Line", NetOffice.WordApi.LineFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Line", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.ThreeD"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ThreeDFormat ThreeD
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ThreeDFormat>(this, "ThreeD", NetOffice.WordApi.ThreeDFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ThreeD", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.TextColor"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.ColorFormat TextColor
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "TextColor", NetOffice.WordApi.ColorFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Ligatures"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdLigatures Ligatures
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLigatures>(this, "Ligatures");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Ligatures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NumberForm"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdNumberForm NumberForm
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNumberForm>(this, "NumberForm");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NumberForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.NumberSpacing"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdNumberSpacing NumberSpacing
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNumberSpacing>(this, "NumberSpacing");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "NumberSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.ContextualAlternates"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public Int32 ContextualAlternates
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ContextualAlternates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ContextualAlternates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.StylisticSet"/> </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Enums.WdStylisticSet StylisticSet
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStylisticSet>(this, "StylisticSet");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "StylisticSet", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Grow"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Grow()
		{
			 Factory.ExecuteMethod(this, "Grow");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Shrink"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Shrink()
		{
			 Factory.ExecuteMethod(this, "Shrink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.Reset"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Font.SetAsTemplateDefault"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetAsTemplateDefault()
		{
			 Factory.ExecuteMethod(this, "SetAsTemplateDefault");
		}

		#endregion

		#pragma warning restore
	}
}
