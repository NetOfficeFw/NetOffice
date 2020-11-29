﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Paragraph 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Paragraph : COMObject
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
                    _type = typeof(Paragraph);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Paragraph(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Paragraph(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraph(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Range"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Range
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Range", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Parent"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Format"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ParagraphFormat Format
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ParagraphFormat>(this, "Format", NetOffice.WordApi.ParagraphFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.TabStops"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TabStops TabStops
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TabStops>(this, "TabStops", NetOffice.WordApi.TabStops.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "TabStops", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Borders"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.DropCap"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.DropCap DropCap
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DropCap>(this, "DropCap", NetOffice.WordApi.DropCap.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Style"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public object Style
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Alignment"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdParagraphAlignment Alignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdParagraphAlignment>(this, "Alignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Alignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.KeepTogether"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 KeepTogether
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "KeepTogether");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeepTogether", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.KeepWithNext"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 KeepWithNext
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "KeepWithNext");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KeepWithNext", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.PageBreakBefore"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 PageBreakBefore
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PageBreakBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageBreakBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.NoLineNumber"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 NoLineNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NoLineNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoLineNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.RightIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single RightIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RightIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.LeftIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LeftIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LeftIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.FirstLineIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single FirstLineIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "FirstLineIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.LineSpacing"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LineSpacing
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LineSpacing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineSpacing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.LineSpacingRule"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLineSpacing LineSpacingRule
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLineSpacing>(this, "LineSpacingRule");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LineSpacingRule", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SpaceBefore"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single SpaceBefore
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SpaceBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpaceBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SpaceAfter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single SpaceAfter
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "SpaceAfter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpaceAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Hyphenation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Hyphenation
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Hyphenation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Hyphenation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.WidowControl"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 WidowControl
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WidowControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WidowControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Shading"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.FarEastLineBreakControl"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 FarEastLineBreakControl
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FarEastLineBreakControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FarEastLineBreakControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.WordWrap"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 WordWrap
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "WordWrap");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WordWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.HangingPunctuation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HangingPunctuation
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HangingPunctuation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HangingPunctuation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.HalfWidthPunctuationOnTopOfLine"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HalfWidthPunctuationOnTopOfLine
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "HalfWidthPunctuationOnTopOfLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HalfWidthPunctuationOnTopOfLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.AddSpaceBetweenFarEastAndAlpha"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 AddSpaceBetweenFarEastAndAlpha
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AddSpaceBetweenFarEastAndAlpha");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddSpaceBetweenFarEastAndAlpha", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.AddSpaceBetweenFarEastAndDigit"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 AddSpaceBetweenFarEastAndDigit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AddSpaceBetweenFarEastAndDigit");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AddSpaceBetweenFarEastAndDigit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.BaseLineAlignment"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdBaselineAlignment BaseLineAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdBaselineAlignment>(this, "BaseLineAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BaseLineAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.AutoAdjustRightIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 AutoAdjustRightIndent
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "AutoAdjustRightIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoAdjustRightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.DisableLineHeightGrid"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 DisableLineHeightGrid
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DisableLineHeightGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableLineHeightGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OutlineLevel"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOutlineLevel OutlineLevel
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOutlineLevel>(this, "OutlineLevel");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OutlineLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.CharacterUnitRightIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single CharacterUnitRightIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CharacterUnitRightIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CharacterUnitRightIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.CharacterUnitLeftIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single CharacterUnitLeftIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CharacterUnitLeftIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CharacterUnitLeftIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.CharacterUnitFirstLineIndent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single CharacterUnitFirstLineIndent
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CharacterUnitFirstLineIndent");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CharacterUnitFirstLineIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.LineUnitBefore"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LineUnitBefore
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LineUnitBefore");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineUnitBefore", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.LineUnitAfter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LineUnitAfter
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LineUnitAfter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LineUnitAfter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ReadingOrder"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdReadingOrder ReadingOrder
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdReadingOrder>(this, "ReadingOrder");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ReadingOrder", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ID"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ID
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ID");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SpaceBeforeAuto"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SpaceBeforeAuto
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SpaceBeforeAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpaceBeforeAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SpaceAfterAuto"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SpaceAfterAuto
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SpaceAfterAuto");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpaceAfterAuto", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.IsStyleSeparator"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool IsStyleSeparator
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsStyleSeparator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.MirrorIndents"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public Int32 MirrorIndents
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MirrorIndents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MirrorIndents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.TextboxTightWrap"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTextboxTightWrap TextboxTightWrap
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTextboxTightWrap>(this, "TextboxTightWrap");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextboxTightWrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListNumberOriginal"/> </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 get_ListNumberOriginal(Int16 level)
		{
			return Factory.ExecuteInt16PropertyGet(this, "ListNumberOriginal", level);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Alias for get_ListNumberOriginal
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListNumberOriginal"/> </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16), Redirect("get_ListNumberOriginal")]
		public Int16 ListNumberOriginal(Int16 level)
		{
			return get_ListNumberOriginal(level);
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ParaID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ParaID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 TextID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.paragraph.collapsedstate"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool CollapsedState
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CollapsedState");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CollapsedState", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.paragraph.collapseheadingbydefault"/> </remarks>
		[SupportByVersion("Word", 15, 16)]
		public bool CollapseHeadingByDefault
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CollapseHeadingByDefault");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CollapseHeadingByDefault", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.CloseUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CloseUp()
		{
			 Factory.ExecuteMethod(this, "CloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OpenUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenUp()
		{
			 Factory.ExecuteMethod(this, "OpenUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OpenOrCloseUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenOrCloseUp()
		{
			 Factory.ExecuteMethod(this, "OpenOrCloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.TabHangingIndent"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TabHangingIndent(Int16 count)
		{
			 Factory.ExecuteMethod(this, "TabHangingIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.TabIndent"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TabIndent(Int16 count)
		{
			 Factory.ExecuteMethod(this, "TabIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Reset"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Space1"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space1()
		{
			 Factory.ExecuteMethod(this, "Space1");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Space15"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space15()
		{
			 Factory.ExecuteMethod(this, "Space15");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Space2"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space2()
		{
			 Factory.ExecuteMethod(this, "Space2");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.IndentCharWidth"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IndentCharWidth(Int16 count)
		{
			 Factory.ExecuteMethod(this, "IndentCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.IndentFirstLineCharWidth"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IndentFirstLineCharWidth(Int16 count)
		{
			 Factory.ExecuteMethod(this, "IndentFirstLineCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Next"/> </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Next(object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Next", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Next"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Next()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Next", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Previous"/> </remarks>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Previous(object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Previous", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Previous"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Previous()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Previous", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OutlinePromote"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlinePromote()
		{
			 Factory.ExecuteMethod(this, "OutlinePromote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OutlineDemote"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlineDemote()
		{
			 Factory.ExecuteMethod(this, "OutlineDemote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.OutlineDemoteToBody"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlineDemoteToBody()
		{
			 Factory.ExecuteMethod(this, "OutlineDemoteToBody");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Indent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Indent()
		{
			 Factory.ExecuteMethod(this, "Indent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.Outdent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Outdent()
		{
			 Factory.ExecuteMethod(this, "Outdent");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SelectNumber"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void SelectNumber()
		{
			 Factory.ExecuteMethod(this, "SelectNumber");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
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
		public void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8, object level9)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7, level8, level9 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo()
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", level1);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", level1, level2);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2, object level3)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", level1, level2, level3);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2, object level3, object level4)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", level1, level2, level3, level4);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
		/// <param name="level1">optional Int16 Level1 = 0</param>
		/// <param name="level2">optional Int16 Level2 = 0</param>
		/// <param name="level3">optional Int16 Level3 = 0</param>
		/// <param name="level4">optional Int16 Level4 = 0</param>
		/// <param name="level5">optional Int16 Level5 = 0</param>
		/// <param name="level6">optional Int16 Level6 = 0</param>
		/// <param name="level7">optional Int16 Level7 = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ListAdvanceTo"/> </remarks>
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
		public void ListAdvanceTo(object level1, object level2, object level3, object level4, object level5, object level6, object level7, object level8)
		{
			 Factory.ExecuteMethod(this, "ListAdvanceTo", new object[]{ level1, level2, level3, level4, level5, level6, level7, level8 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.ResetAdvanceTo"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ResetAdvanceTo()
		{
			 Factory.ExecuteMethod(this, "ResetAdvanceTo");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.SeparateList"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void SeparateList()
		{
			 Factory.ExecuteMethod(this, "SeparateList");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraph.JoinList"/> </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void JoinList()
		{
			 Factory.ExecuteMethod(this, "JoinList");
		}

		#endregion

		#pragma warning restore
	}
}
