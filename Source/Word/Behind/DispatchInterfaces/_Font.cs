using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface _Font 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _Font : COMObject, NetOffice.WordApi._Font
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
                    _contractType = typeof(NetOffice.WordApi._Font);
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
                    _type = typeof(_Font);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _Font() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198225.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197420.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822380.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Font Duplicate
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Font>(this, "Duplicate", typeof(NetOffice.WordApi.Font));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822177.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Bold
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Bold");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836352.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Italic
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Italic");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835996.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Hidden
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hidden");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840430.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SmallCaps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SmallCaps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SmallCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 AllCaps
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AllCaps");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AllCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837930.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 StrikeThrough
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "StrikeThrough");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "StrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845854.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 DoubleStrikeThrough
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DoubleStrikeThrough");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DoubleStrikeThrough", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837918.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdColorIndex ColorIndex
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "ColorIndex");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195668.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Subscript
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Subscript");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Subscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845184.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Superscript
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Superscript");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Superscript", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194284.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdUnderline Underline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdUnderline>(this, "Underline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834545.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Size
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Size");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Size", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195908.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193398.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Position
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Position");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Position", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195646.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Spacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Spacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Spacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195925.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Scaling
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Scaling");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Scaling", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839784.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Shadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Shadow");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Shadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840442.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Outline
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Outline");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Outline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836373.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Emboss
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Emboss");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Emboss", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192149.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single Kerning
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "Kerning");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Kerning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835148.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Engrave
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Engrave");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Engrave", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdAnimation Animation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdAnimation>(this, "Animation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Animation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835510.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Borders Borders
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Borders>(this, "Borders", typeof(NetOffice.WordApi.Borders));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834890.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Shading Shading
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Shading>(this, "Shading", typeof(NetOffice.WordApi.Shading));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191793.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEmphasisMark>(this, "EmphasisMark");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "EmphasisMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837018.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool DisableCharacterSpaceGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisableCharacterSpaceGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableCharacterSpaceGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822977.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NameFarEast
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameFarEast");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameFarEast", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834607.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NameAscii
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameAscii");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameAscii", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838727.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NameOther
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameOther");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameOther", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdColor Color
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "Color");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821059.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 BoldBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "BoldBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "BoldBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196202.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 ItalicBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ItalicBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ItalicBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195703.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single SizeBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SizeBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SizeBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194616.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string NameBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "NameBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NameBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838272.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdColorIndex ColorIndexBi
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "ColorIndexBi");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ColorIndexBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838275.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdColor DiacriticColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "DiacriticColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DiacriticColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840299.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdColor UnderlineColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColor>(this, "UnderlineColor");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "UnderlineColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836301.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.GlowFormat Glow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.GlowFormat>(this, "Glow", typeof(NetOffice.WordApi.GlowFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Glow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835746.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.ReflectionFormat Reflection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReflectionFormat>(this, "Reflection", typeof(NetOffice.WordApi.ReflectionFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Reflection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194468.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.ShadowFormat TextShadow
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShadowFormat>(this, "TextShadow", typeof(NetOffice.WordApi.ShadowFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "TextShadow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837243.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.FillFormat Fill
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FillFormat>(this, "Fill", typeof(NetOffice.WordApi.FillFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Fill", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193611.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.LineFormat Line
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LineFormat>(this, "Line", typeof(NetOffice.WordApi.LineFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Line", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836007.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.ThreeDFormat ThreeD
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ThreeDFormat>(this, "ThreeD", typeof(NetOffice.WordApi.ThreeDFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ThreeD", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194246.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.ColorFormat TextColor
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ColorFormat>(this, "TextColor", typeof(NetOffice.WordApi.ColorFormat));
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840369.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLigatures Ligatures
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLigatures>(this, "Ligatures");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Ligatures", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820804.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdNumberForm NumberForm
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNumberForm>(this, "NumberForm");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "NumberForm", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194804.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdNumberSpacing NumberSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdNumberSpacing>(this, "NumberSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "NumberSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845101.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual Int32 ContextualAlternates
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ContextualAlternates");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ContextualAlternates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839787.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdStylisticSet StylisticSet
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStylisticSet>(this, "StylisticSet");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "StylisticSet", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845371.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Grow()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Grow");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197751.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Shrink()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Shrink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Reset()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834270.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SetAsTemplateDefault()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAsTemplateDefault");
		}

		#endregion

		#pragma warning restore
	}
}


