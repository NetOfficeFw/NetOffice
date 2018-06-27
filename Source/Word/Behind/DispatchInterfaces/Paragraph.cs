using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Paragraph 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845307.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Paragraph : COMObject, NetOffice.WordApi.Paragraph
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
                    _contractType = typeof(NetOffice.WordApi.Paragraph);
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
                    _type = typeof(Paragraph);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Paragraph() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198036.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Range Range
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", typeof(NetOffice.WordApi.Range));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193407.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840156.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845450.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838951.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ParagraphFormat Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ParagraphFormat>(this, "Format", typeof(NetOffice.WordApi.ParagraphFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839415.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.TabStops TabStops
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TabStops>(this, "TabStops", typeof(NetOffice.WordApi.TabStops));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "TabStops", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834531.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193363.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.DropCap DropCap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DropCap>(this, "DropCap", typeof(NetOffice.WordApi.DropCap));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835808.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual object Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844837.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdParagraphAlignment Alignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdParagraphAlignment>(this, "Alignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Alignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835246.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 KeepTogether
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "KeepTogether");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeepTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196384.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 KeepWithNext
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "KeepWithNext");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "KeepWithNext", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821677.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 PageBreakBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PageBreakBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PageBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840640.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 NoLineNumber
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NoLineNumber");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoLineNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192362.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single RightIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "RightIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192031.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single LeftIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LeftIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194707.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single FirstLineIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "FirstLineIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single LineSpacing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LineSpacing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844903.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLineSpacing LineSpacingRule
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineSpacing>(this, "LineSpacingRule");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LineSpacingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194228.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single SpaceBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SpaceBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191990.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single SpaceAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "SpaceAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834839.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Hyphenation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Hyphenation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Hyphenation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196595.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 WidowControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WidowControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WidowControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822389.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834812.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 FarEastLineBreakControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "FarEastLineBreakControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FarEastLineBreakControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 WordWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "WordWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "WordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822614.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 HangingPunctuation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HangingPunctuation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HangingPunctuation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196345.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 HalfWidthPunctuationOnTopOfLine
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HalfWidthPunctuationOnTopOfLine");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HalfWidthPunctuationOnTopOfLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194017.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 AddSpaceBetweenFarEastAndAlpha
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AddSpaceBetweenFarEastAndAlpha");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddSpaceBetweenFarEastAndAlpha", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836922.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 AddSpaceBetweenFarEastAndDigit
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AddSpaceBetweenFarEastAndDigit");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AddSpaceBetweenFarEastAndDigit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192567.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdBaselineAlignment BaseLineAlignment
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdBaselineAlignment>(this, "BaseLineAlignment");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BaseLineAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192561.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 AutoAdjustRightIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "AutoAdjustRightIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoAdjustRightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821565.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 DisableLineHeightGrid
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DisableLineHeightGrid");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisableLineHeightGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197419.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdOutlineLevel OutlineLevel
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOutlineLevel>(this, "OutlineLevel");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OutlineLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840651.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single CharacterUnitRightIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "CharacterUnitRightIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CharacterUnitRightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192004.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single CharacterUnitLeftIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "CharacterUnitLeftIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CharacterUnitLeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844832.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single CharacterUnitFirstLineIndent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "CharacterUnitFirstLineIndent");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CharacterUnitFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839860.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single LineUnitBefore
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LineUnitBefore");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineUnitBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845221.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Single LineUnitAfter
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteSinglePropertyGet(this, "LineUnitAfter");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "LineUnitAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836291.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdReadingOrder ReadingOrder
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdReadingOrder>(this, "ReadingOrder");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ReadingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840491.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string ID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195344.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SpaceBeforeAuto
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SpaceBeforeAuto");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceBeforeAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838152.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 SpaceAfterAuto
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SpaceAfterAuto");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SpaceAfterAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198321.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual bool IsStyleSeparator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsStyleSeparator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196956.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual Int32 MirrorIndents
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MirrorIndents");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MirrorIndents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822871.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdTextboxTightWrap TextboxTightWrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTextboxTightWrap>(this, "TextboxTightWrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "TextboxTightWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840388.aspx </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int16 get_ListNumberOriginal(Int16 level)
		{
			return InvokerService.InvokeInternal.ExecuteInt16PropertyGet(this, "ListNumberOriginal", level);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Alias for get_ListNumberOriginal
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840388.aspx </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16), Redirect("get_ListNumberOriginal")]
		public virtual Int16 ListNumberOriginal(Int16 level)
		{
			return get_ListNumberOriginal(level);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 ParaID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ParaID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public virtual Int32 TextID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "TextID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229749.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual bool CollapsedState
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CollapsedState");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CollapsedState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232325.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public virtual bool CollapseHeadingByDefault
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CollapseHeadingByDefault");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CollapseHeadingByDefault", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839929.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void CloseUp()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197449.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenUp()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836114.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OpenOrCloseUp()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenOrCloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837340.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void TabHangingIndent(Int16 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TabHangingIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198348.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void TabIndent(Int16 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "TabIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834877.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Reset()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839871.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Space1()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Space1");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838089.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Space15()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Space15");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Space2()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Space2");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839146.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void IndentCharWidth(Int16 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "IndentCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191696.aspx </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void IndentFirstLineCharWidth(Int16 count)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "IndentFirstLineCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196530.aspx </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Paragraph Next(object count)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Next", typeof(NetOffice.WordApi.Paragraph), count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196530.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Paragraph Next()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Next", typeof(NetOffice.WordApi.Paragraph));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845423.aspx </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Paragraph Previous(object count)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Previous", typeof(NetOffice.WordApi.Paragraph), count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845423.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Paragraph Previous()
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Previous", typeof(NetOffice.WordApi.Paragraph));
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821068.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OutlinePromote()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlinePromote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844908.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OutlineDemote()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineDemote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194241.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void OutlineDemoteToBody()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OutlineDemoteToBody");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196925.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Indent()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Indent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192199.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void Outdent()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "Outdent");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834939.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public virtual void SelectNumber()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SelectNumber");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		/// <param name="level8">optional Int16 Level8 = 0</param>
		/// <param name="level9">optional Int16 Level9 = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8, object level9)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7, level8, level9 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", level1);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", level1, level2);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", level1, level2, level3);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", level1, level2, level3, level4);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194432.aspx </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		/// <param name="level8">optional Int16 Level8 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7, level8 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196834.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void ResetAdvanceTo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ResetAdvanceTo");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845535.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void SeparateList()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SeparateList");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192823.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual void JoinList()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "JoinList");
		}

		#endregion

		#pragma warning restore
	}
}


