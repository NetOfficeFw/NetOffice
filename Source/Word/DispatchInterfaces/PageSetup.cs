using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface PageSetup 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191765.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195084.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196544.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844992.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838093.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192545.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822398.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836129.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839979.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197123.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840372.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821149.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196971.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839341.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838676.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836404.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841028.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845397.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193079.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821984.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195626.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837515.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836298.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822350.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845091.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838147.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820990.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839364.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230012.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838475.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837728.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834942.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198290.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845067.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194058.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822568.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191756.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TogglePortrait()
		{
			 Factory.ExecuteMethod(this, "TogglePortrait");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193732.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetAsTemplateDefault()
		{
			 Factory.ExecuteMethod(this, "SetAsTemplateDefault");
		}

		#endregion

		#pragma warning restore
	}
}
