using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// PageSetup
	/// </summary>
	[SyntaxBypass]
 	public class PageSetup_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PageSetup_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PageSetup_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageSetup_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <param name="index">optional object index</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_PrintQuality(object index)
		{
			return Factory.ExecuteVariantPropertyGet(this, "PrintQuality", index);
		}

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index">optional object index</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_PrintQuality(object index, object value)
		{
			Factory.ExecutePropertySet(this, "PrintQuality", index, value);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_PrintQuality
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx </remarks>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), Redirect("get_PrintQuality")]
		public object PrintQuality(object index)
		{
			return get_PrintQuality(index);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface PageSetup 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196103.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class PageSetup : PageSetup_
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
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836795.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835290.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836825.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834921.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool BlackAndWhite
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BlackAndWhite");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BlackAndWhite", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198024.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double BottomMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "BottomMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BottomMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839224.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string CenterFooter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CenterFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CenterFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835588.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string CenterHeader
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CenterHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CenterHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822498.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CenterHorizontally
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CenterHorizontally");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CenterHorizontally", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841275.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool CenterVertically
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CenterVertically");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CenterVertically", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlObjectSize ChartSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlObjectSize>(this, "ChartSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ChartSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194155.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool Draft
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Draft");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Draft", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821636.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 FirstPageNumber
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "FirstPageNumber");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FirstPageNumber", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194582.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FitToPagesTall
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FitToPagesTall");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FitToPagesTall", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194365.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object FitToPagesWide
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FitToPagesWide");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FitToPagesWide", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838218.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double FooterMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "FooterMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FooterMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838659.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double HeaderMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "HeaderMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HeaderMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822322.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string LeftFooter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LeftFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836805.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string LeftHeader
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "LeftHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821496.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double LeftMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "LeftMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LeftMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839060.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlOrder Order
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlOrder>(this, "Order");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Order", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836834.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPageOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPageOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834612.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPaperSize PaperSize
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPaperSize>(this, "PaperSize");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PaperSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839804.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PrintArea
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrintArea");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintArea", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836524.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PrintGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192995.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PrintHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822102.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public bool PrintNotes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "PrintNotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintNotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194768.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object PrintQuality
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PrintQuality");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PrintQuality", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835214.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PrintTitleColumns
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrintTitleColumns");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintTitleColumns", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839961.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string PrintTitleRows
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "PrintTitleRows");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "PrintTitleRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840096.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string RightFooter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RightFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836463.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public string RightHeader
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RightHeader");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightHeader", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836771.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double RightMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "RightMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RightMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194770.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Double TopMargin
		{
			get
			{
				return Factory.ExecuteDoublePropertyGet(this, "TopMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TopMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197028.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public object Zoom
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Zoom");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Zoom", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194941.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPrintLocation PrintComments
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintLocation>(this, "PrintComments");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197865.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPrintErrors PrintErrors
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPrintErrors>(this, "PrintErrors");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PrintErrors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838811.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic CenterHeaderPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterHeaderPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822635.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic CenterFooterPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "CenterFooterPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194861.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic LeftHeaderPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftHeaderPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195659.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic LeftFooterPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "LeftFooterPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196670.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic RightHeaderPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightHeaderPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840868.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Graphic RightFooterPicture
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Graphic>(this, "RightFooterPicture", NetOffice.ExcelApi.Graphic.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839665.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool OddAndEvenPagesHeaderFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OddAndEvenPagesHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OddAndEvenPagesHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838040.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool DifferentFirstPageHeaderFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DifferentFirstPageHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DifferentFirstPageHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838420.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool ScaleWithDocHeaderFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ScaleWithDocHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ScaleWithDocHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838197.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public bool AlignMarginsHeaderFooter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AlignMarginsHeaderFooter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AlignMarginsHeaderFooter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840030.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Pages Pages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Pages>(this, "Pages", NetOffice.ExcelApi.Pages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839989.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Page EvenPage
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "EvenPage", NetOffice.ExcelApi.Page.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835921.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Page FirstPage
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Page>(this, "FirstPage", NetOffice.ExcelApi.Page.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}
