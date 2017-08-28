using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface View 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822898.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class View : COMObject
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
                    _type = typeof(View);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public View(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public View(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public View(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197224.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821357.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844900.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844848.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdViewType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdViewType>(this, "Type");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Type", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840409.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool FullScreen
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FullScreen");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FullScreen", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834875.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192196.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowAll
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAll");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAll", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840713.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowFieldCodes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFieldCodes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFieldCodes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192217.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MailMergeDataView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MailMergeDataView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MailMergeDataView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835476.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Magnifier
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Magnifier");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Magnifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191722.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowFirstLineOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFirstLineOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFirstLineOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821926.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowFormat
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFormat");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196289.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Zoom Zoom
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Zoom>(this, "Zoom", NetOffice.WordApi.Zoom.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197841.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowObjectAnchors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowObjectAnchors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowObjectAnchors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836054.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowTextBoundaries
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTextBoundaries");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTextBoundaries", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839975.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowHighlight
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHighlight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHighlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840802.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowDrawings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowDrawings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowDrawings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840032.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowTabs
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowTabs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowTabs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837933.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowSpaces
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowSpaces");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowSpaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191746.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowParagraphs
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowParagraphs");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowParagraphs", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192337.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowHyphens
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHyphens");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHyphens", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839598.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowHiddenText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowHiddenText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowHiddenText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840568.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool WrapToWindow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "WrapToWindow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "WrapToWindow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197748.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowPicturePlaceHolders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowPicturePlaceHolders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowPicturePlaceHolders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192158.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowBookmarks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowBookmarks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowBookmarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195455.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFieldShading FieldShading
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFieldShading>(this, "FieldShading");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "FieldShading", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowAnimation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowAnimation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowAnimation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844911.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool TableGridlines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TableGridlines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TableGridlines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EnlargeFontsLessThan
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "EnlargeFontsLessThan");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EnlargeFontsLessThan", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845503.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowMainTextLayer
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowMainTextLayer");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMainTextLayer", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834537.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSeekView SeekView
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSeekView>(this, "SeekView");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SeekView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196649.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSpecialPane SplitSpecial
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSpecialPane>(this, "SplitSpecial");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SplitSpecial", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BrowseToWindow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BrowseToWindow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BrowseToWindow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839824.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ShowOptionalBreaks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowOptionalBreaks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowOptionalBreaks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197553.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisplayPageBoundaries
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayPageBoundaries");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayPageBoundaries", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool DisplaySmartTags
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplaySmartTags");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplaySmartTags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836994.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowRevisionsAndComments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowRevisionsAndComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowRevisionsAndComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844875.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowComments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193645.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowInsertionsAndDeletions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowInsertionsAndDeletions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowInsertionsAndDeletions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839560.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool ShowFormatChanges
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowFormatChanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowFormatChanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsView RevisionsView
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsView>(this, "RevisionsView");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisionsView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsMode RevisionsMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsMode>(this, "RevisionsMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisionsMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840523.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public Single RevisionsBalloonWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "RevisionsBalloonWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RevisionsBalloonWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840447.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsBalloonWidthType RevisionsBalloonWidthType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsBalloonWidthType>(this, "RevisionsBalloonWidthType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisionsBalloonWidthType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197146.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsBalloonMargin RevisionsBalloonSide
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsBalloonMargin>(this, "RevisionsBalloonSide");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RevisionsBalloonSide", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Reviewers Reviewers
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Reviewers>(this, "Reviewers", NetOffice.WordApi.Reviewers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821251.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool RevisionsBalloonShowConnectingLines
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RevisionsBalloonShowConnectingLines");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RevisionsBalloonShowConnectingLines", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839617.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ReadingLayout
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadingLayout");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayout", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198229.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 ShowXMLMarkup
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ShowXMLMarkup");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowXMLMarkup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840275.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public Int32 ShadeEditableRanges
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ShadeEditableRanges");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShadeEditableRanges", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196812.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ShowInkAnnotations
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowInkAnnotations");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowInkAnnotations", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197827.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool DisplayBackgrounds
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayBackgrounds");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayBackgrounds", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197998.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ReadingLayoutActualView
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadingLayoutActualView");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayoutActualView", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public bool ReadingLayoutAllowMultiplePages
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadingLayoutAllowMultiplePages");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayoutAllowMultiplePages", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ReadingLayoutAllowEditing
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadingLayoutAllowEditing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ReadingLayoutAllowEditing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845085.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdReadingLayoutMargin ReadingLayoutTruncateMargins
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdReadingLayoutMargin>(this, "ReadingLayoutTruncateMargins");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ReadingLayoutTruncateMargins", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194113.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowMarkupAreaHighlight
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowMarkupAreaHighlight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowMarkupAreaHighlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844988.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Panning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Panning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Panning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837472.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ShowCropMarks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowCropMarks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowCropMarks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192820.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.Enums.WdRevisionsMode MarkupMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdRevisionsMode>(this, "MarkupMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MarkupMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839848.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool ConflictMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ConflictMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConflictMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845837.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public bool ShowOtherAuthors
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ShowOtherAuthors");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ShowOtherAuthors", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231107.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.RevisionsFilter RevisionsFilter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.RevisionsFilter>(this, "RevisionsFilter", NetOffice.WordApi.RevisionsFilter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231031.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdPageColor PageColor
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdPageColor>(this, "PageColor");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PageColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230868.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public NetOffice.WordApi.Enums.WdColumnWidth ColumnWidth
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColumnWidth>(this, "ColumnWidth");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ColumnWidth", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836702.aspx </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CollapseOutline(object range)
		{
			 Factory.ExecuteMethod(this, "CollapseOutline", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836702.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CollapseOutline()
		{
			 Factory.ExecuteMethod(this, "CollapseOutline");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194789.aspx </remarks>
		/// <param name="range">optional object range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ExpandOutline(object range)
		{
			 Factory.ExecuteMethod(this, "ExpandOutline", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194789.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ExpandOutline()
		{
			 Factory.ExecuteMethod(this, "ExpandOutline");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192618.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ShowAllHeadings()
		{
			 Factory.ExecuteMethod(this, "ShowAllHeadings");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836914.aspx </remarks>
		/// <param name="level">Int32 level</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ShowHeading(Int32 level)
		{
			 Factory.ExecuteMethod(this, "ShowHeading", level);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841033.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PreviousHeaderFooter()
		{
			 Factory.ExecuteMethod(this, "PreviousHeaderFooter");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195041.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void NextHeaderFooter()
		{
			 Factory.ExecuteMethod(this, "NextHeaderFooter");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231576.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void ExpandAllHeadings()
		{
			 Factory.ExecuteMethod(this, "ExpandAllHeadings");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227347.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public void CollapseAllHeadings()
		{
			 Factory.ExecuteMethod(this, "CollapseAllHeadings");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[SupportByVersion("Word", 15, 16)]
		public void ForceOffscreenUpdate()
		{
			 Factory.ExecuteMethod(this, "ForceOffscreenUpdate");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		[SupportByVersion("Word", 15, 16)]
		public void ForceLowresUpdate()
		{
			 Factory.ExecuteMethod(this, "ForceLowresUpdate");
		}

		#endregion

		#pragma warning restore
	}
}
