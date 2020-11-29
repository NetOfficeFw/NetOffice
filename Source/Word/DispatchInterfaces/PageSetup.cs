﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface PageSetup 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup"/> </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PageSetup : COMObject
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
                    _type = typeof(PageSetup);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PageSetup(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PageSetup(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.TopMargin"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single TopMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "TopMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.BottomMargin"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single BottomMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "BottomMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.LeftMargin"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LeftMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LeftMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.RightMargin"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single RightMargin
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RightMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.Gutter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Gutter
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "Gutter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Gutter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.PageWidth"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PageWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PageWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.PageHeight"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single PageHeight
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "PageHeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PageHeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.Orientation"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.FirstPageTray"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperTray FirstPageTray
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "FirstPageTray");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FirstPageTray", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.OtherPagesTray"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperTray OtherPagesTray
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperTray>(this, "OtherPagesTray");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OtherPagesTray", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.VerticalAlignment"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdVerticalAlignment VerticalAlignment
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdVerticalAlignment>(this, "VerticalAlignment");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "VerticalAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.MirrorMargins"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MirrorMargins
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MirrorMargins");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MirrorMargins", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.HeaderDistance"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single HeaderDistance
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "HeaderDistance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeaderDistance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.FooterDistance"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single FooterDistance
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "FooterDistance");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FooterDistance", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.SectionStart"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSectionStart SectionStart
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSectionStart>(this, "SectionStart");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SectionStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.OddAndEvenPagesHeaderFooter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 OddAndEvenPagesHeaderFooter
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "OddAndEvenPagesHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OddAndEvenPagesHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.DifferentFirstPageHeaderFooter"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 DifferentFirstPageHeaderFooter
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DifferentFirstPageHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DifferentFirstPageHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.SuppressEndnotes"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 SuppressEndnotes
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SuppressEndnotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuppressEndnotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.LineNumbering"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.LineNumbering LineNumbering
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.LineNumbering>(this, "LineNumbering", NetOffice.WordApi.LineNumbering.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "LineNumbering", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.TextColumns"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TextColumns TextColumns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextColumns>(this, "TextColumns", NetOffice.WordApi.TextColumns.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "TextColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.PaperSize"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdPaperSize PaperSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPaperSize>(this, "PaperSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PaperSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.TwoPagesOnOne"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool TwoPagesOnOne
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TwoPagesOnOne");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TwoPagesOnOne", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool GutterOnTop
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GutterOnTop");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GutterOnTop", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.CharsLine"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single CharsLine
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "CharsLine");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CharsLine", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.LinesPage"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single LinesPage
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "LinesPage");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LinesPage", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.pagesetup.showgrid"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.GutterStyle"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdGutterStyleOld GutterStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdGutterStyleOld>(this, "GutterStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GutterStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.SectionDirection"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSectionDirection SectionDirection
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSectionDirection>(this, "SectionDirection");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SectionDirection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.LayoutMode"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLayoutMode LayoutMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLayoutMode>(this, "LayoutMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LayoutMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.GutterPos"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdGutterStyle GutterPos
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdGutterStyle>(this, "GutterPos");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "GutterPos", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.BookFoldPrinting"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool BookFoldPrinting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BookFoldPrinting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BookFoldPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.BookFoldRevPrinting"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool BookFoldRevPrinting
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BookFoldRevPrinting");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BookFoldRevPrinting", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.BookFoldPrintingSheets"/> </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Int32 BookFoldPrintingSheets
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BookFoldPrintingSheets");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BookFoldPrintingSheets", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.TogglePortrait"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TogglePortrait()
		{
			 Factory.ExecuteMethod(this, "TogglePortrait");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Word.PageSetup.SetAsTemplateDefault"/> </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetAsTemplateDefault()
		{
			 Factory.ExecuteMethod(this, "SetAsTemplateDefault");
		}

		#endregion

		#pragma warning restore
	}
}
