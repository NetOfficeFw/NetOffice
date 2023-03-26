using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Paragraphs 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.paragraphs"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class Paragraphs : COMObject, IEnumerableProvider<NetOffice.WordApi.Paragraph>
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
                    _type = typeof(Paragraphs);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Paragraphs(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Paragraphs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Paragraphs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Count"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.First"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph First
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraph>(this, "First", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Last"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Last
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraph>(this, "Last", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Format"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.TabStops"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Borders"/> </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Style"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Alignment"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.KeepTogether"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.KeepWithNext"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.PageBreakBefore"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.NoLineNumber"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.RightIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.LeftIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.FirstLineIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.LineSpacing"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.LineSpacingRule"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.SpaceBefore"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.SpaceAfter"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Hyphenation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.WidowControl"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Shading"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.FarEastLineBreakControl"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.WordWrap"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.HangingPunctuation"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.HalfWidthPunctuationOnTopOfLine"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.AddSpaceBetweenFarEastAndAlpha"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.AddSpaceBetweenFarEastAndDigit"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.BaseLineAlignment"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.AutoAdjustRightIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.DisableLineHeightGrid"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OutlineLevel"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.CharacterUnitRightIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.CharacterUnitLeftIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.CharacterUnitFirstLineIndent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.LineUnitBefore"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.LineUnitAfter"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.ReadingOrder"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.SpaceBeforeAuto"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.SpaceAfterAuto"/> </remarks>
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">Int32 index</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.WordApi.Paragraph this[Int32 index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Item", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Add"/> </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Add(object range)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Add", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Add"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraph Add()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Paragraph>(this, "Add", NetOffice.WordApi.Paragraph.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.CloseUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CloseUp()
		{
			 Factory.ExecuteMethod(this, "CloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OpenUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenUp()
		{
			 Factory.ExecuteMethod(this, "OpenUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OpenOrCloseUp"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OpenOrCloseUp()
		{
			 Factory.ExecuteMethod(this, "OpenOrCloseUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.TabHangingIndent"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TabHangingIndent(Int16 count)
		{
			 Factory.ExecuteMethod(this, "TabHangingIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.TabIndent"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TabIndent(Int16 count)
		{
			 Factory.ExecuteMethod(this, "TabIndent", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Reset"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Space1"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space1()
		{
			 Factory.ExecuteMethod(this, "Space1");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Space15"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space15()
		{
			 Factory.ExecuteMethod(this, "Space15");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Space2"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Space2()
		{
			 Factory.ExecuteMethod(this, "Space2");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.IndentCharWidth"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IndentCharWidth(Int16 count)
		{
			 Factory.ExecuteMethod(this, "IndentCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.IndentFirstLineCharWidth"/> </remarks>
		/// <param name="count">Int16 count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void IndentFirstLineCharWidth(Int16 count)
		{
			 Factory.ExecuteMethod(this, "IndentFirstLineCharWidth", count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OutlinePromote"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlinePromote()
		{
			 Factory.ExecuteMethod(this, "OutlinePromote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OutlineDemote"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlineDemote()
		{
			 Factory.ExecuteMethod(this, "OutlineDemote");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.OutlineDemoteToBody"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void OutlineDemoteToBody()
		{
			 Factory.ExecuteMethod(this, "OutlineDemoteToBody");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Indent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Indent()
		{
			 Factory.ExecuteMethod(this, "Indent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.Outdent"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Outdent()
		{
			 Factory.ExecuteMethod(this, "Outdent");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.IncreaseSpacing"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void IncreaseSpacing()
		{
			 Factory.ExecuteMethod(this, "IncreaseSpacing");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.Paragraphs.DecreaseSpacing"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void DecreaseSpacing()
		{
			 Factory.ExecuteMethod(this, "DecreaseSpacing");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.WordApi.Paragraph>

        ICOMObject IEnumerableProvider<NetOffice.WordApi.Paragraph>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.WordApi.Paragraph>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.WordApi.Paragraph>

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.WordApi.Paragraph> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.WordApi.Paragraph item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Word, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Word", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}