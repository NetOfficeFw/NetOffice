using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface EmailOptions 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194477.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class EmailOptions : COMObject
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
                    _type = typeof(EmailOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public EmailOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public EmailOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public EmailOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197525.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191741.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840769.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839521.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool UseThemeStyle
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseThemeStyle");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseThemeStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840320.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string MarkCommentsWith
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MarkCommentsWith");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkCommentsWith", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821277.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MarkComments
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MarkComments");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MarkComments", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822191.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.EmailSignature EmailSignature
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EmailSignature>(this, "EmailSignature", NetOffice.WordApi.EmailSignature.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845385.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Style ComposeStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "ComposeStyle", NetOffice.WordApi.Style.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836350.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Style ReplyStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "ReplyStyle", NetOffice.WordApi.Style.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840027.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string ThemeName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ThemeName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ThemeName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy1
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dummy1");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Dummy2
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Dummy2");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840679.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool NewColorOnReply
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NewColorOnReply");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NewColorOnReply", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839519.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Style PlainTextStyle
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Style>(this, "PlainTextStyle", NetOffice.WordApi.Style.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845433.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool UseThemeStyleOnReply
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseThemeStyleOnReply");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseThemeStyleOnReply", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192628.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyHeadings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyHeadings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyHeadings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192399.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyBorders
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyBorders");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyBorders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837236.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyBulletedLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyBulletedLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyBulletedLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193750.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyNumberedLists
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyNumberedLists");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyNumberedLists", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193405.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceQuotes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceQuotes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceQuotes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838943.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceSymbols
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceSymbols");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceSymbols", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837891.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceOrdinals
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceOrdinals");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceOrdinals", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194415.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceFractions
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceFractions");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceFractions", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193085.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplacePlainTextEmphasis
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplacePlainTextEmphasis");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplacePlainTextEmphasis", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837035.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeFormatListItemBeginning
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeFormatListItemBeginning");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeFormatListItemBeginning", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840029.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeDefineStyles
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeDefineStyles");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeDefineStyles", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823235.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceHyperlinks
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceHyperlinks");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceHyperlinks", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839729.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyTables
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyTables");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyTables", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835434.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyFirstIndents
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyFirstIndents");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyFirstIndents", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839691.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyDates
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyDates");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyDates", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837002.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeApplyClosings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeApplyClosings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeApplyClosings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837502.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeMatchParentheses
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeMatchParentheses");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeMatchParentheses", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845297.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeReplaceFarEastDashes
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeReplaceFarEastDashes");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeReplaceFarEastDashes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838514.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeDeleteAutoSpaces
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeDeleteAutoSpaces");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeDeleteAutoSpaces", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840277.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeInsertClosings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeInsertClosings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeInsertClosings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193845.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeAutoLetterWizard
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeAutoLetterWizard");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeAutoLetterWizard", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845410.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool AutoFormatAsYouTypeInsertOvers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoFormatAsYouTypeInsertOvers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoFormatAsYouTypeInsertOvers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195336.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool RelyOnCSS
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RelyOnCSS");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RelyOnCSS", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195303.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEmailHTMLFidelity HTMLFidelity
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEmailHTMLFidelity>(this, "HTMLFidelity");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HTMLFidelity", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool EmbedSmartTag
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmbedSmartTag");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmbedSmartTag", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195043.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool TabIndentKey
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TabIndentKey");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TabIndentKey", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Dummy3()
		{
			 Factory.ExecuteMethod(this, "Dummy3");
		}

		#endregion

		#pragma warning restore
	}
}
