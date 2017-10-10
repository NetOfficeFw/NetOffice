using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// Selection
	/// </summary>
	[SyntaxBypass]
 	public class Selection_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Selection_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Selection_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="dataOnly">optional bool dataOnly</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx
		[SupportByVersion("Word", 11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string get_XML(object dataOnly)
		{
			return Factory.ExecuteStringPropertyGet(this, "XML", dataOnly);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Alias for get_XML
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
		/// <param name="dataOnly">optional bool dataOnly</param>
		[SupportByVersion("Word", 11,12,14,15,16), Redirect("get_XML")]
		public string XML(object dataOnly)
		{
			return get_XML(dataOnly);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface Selection 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821411.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Selection : Selection_
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
                    _type = typeof(Selection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Selection(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Selection(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Selection(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192754.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836670.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range FormattedText
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "FormattedText", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "FormattedText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839485.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Start
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Start");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Start", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834869.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 End
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "End");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "End", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837859.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Font>(this, "Font", NetOffice.WordApi.Font.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821048.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSelectionType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSelectionType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191739.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdStoryType StoryType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdStoryType>(this, "StoryType");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838978.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845908.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables Tables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "Tables", NetOffice.WordApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837460.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Words Words
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Words>(this, "Words", NetOffice.WordApi.Words.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193720.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sentences Sentences
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sentences>(this, "Sentences", NetOffice.WordApi.Sentences.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196946.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Characters Characters
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Characters>(this, "Characters", NetOffice.WordApi.Characters.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197009.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Footnotes Footnotes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Footnotes>(this, "Footnotes", NetOffice.WordApi.Footnotes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841006.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Endnotes Endnotes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Endnotes>(this, "Endnotes", NetOffice.WordApi.Endnotes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823219.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Comments Comments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Comments>(this, "Comments", NetOffice.WordApi.Comments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195296.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Cells Cells
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Cells>(this, "Cells", NetOffice.WordApi.Cells.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836277.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Sections Sections
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Sections>(this, "Sections", NetOffice.WordApi.Sections.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840393.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Paragraphs Paragraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Paragraphs>(this, "Paragraphs", NetOffice.WordApi.Paragraphs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193012.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192021.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845839.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Fields Fields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Fields>(this, "Fields", NetOffice.WordApi.Fields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838906.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.FormFields FormFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FormFields>(this, "FormFields", NetOffice.WordApi.FormFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838307.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frames Frames
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frames>(this, "Frames", NetOffice.WordApi.Frames.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193858.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ParagraphFormat>(this, "ParagraphFormat", NetOffice.WordApi.ParagraphFormat.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ParagraphFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197430.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.PageSetup PageSetup
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PageSetup>(this, "PageSetup", NetOffice.WordApi.PageSetup.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "PageSetup", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193356.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Bookmarks Bookmarks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Bookmarks>(this, "Bookmarks", NetOffice.WordApi.Bookmarks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836357.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StoryLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "StoryLength");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838983.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageID
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageID");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196398.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDFarEast");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LanguageIDFarEast", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191830.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDOther");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "LanguageIDOther", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838134.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Hyperlinks Hyperlinks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Hyperlinks>(this, "Hyperlinks", NetOffice.WordApi.Hyperlinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194663.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Columns Columns
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Columns>(this, "Columns", NetOffice.WordApi.Columns.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821842.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Rows Rows
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Rows>(this, "Rows", NetOffice.WordApi.Rows.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836744.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.HeaderFooter HeaderFooter
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HeaderFooter>(this, "HeaderFooter", NetOffice.WordApi.HeaderFooter.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845161.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsEndOfRowMark
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsEndOfRowMark");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840519.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BookmarkID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BookmarkID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193388.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 PreviousBookmarkID
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "PreviousBookmarkID");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197434.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Find Find
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Find>(this, "Find", NetOffice.WordApi.Find.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845594.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Information(NetOffice.WordApi.Enums.WdInformation type)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Information", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_Information
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820800.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_Information")]
		public object Information(NetOffice.WordApi.Enums.WdInformation type)
		{
			return get_Information(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837479.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdSelectionFlags Flags
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdSelectionFlags>(this, "Flags");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Flags", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835497.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Active
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Active");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820824.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool StartIsActive
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "StartIsActive");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "StartIsActive", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822970.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IPAtEndOfLine
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IPAtEndOfLine");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821400.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExtendMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ExtendMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ExtendMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839310.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ColumnSelectMode
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ColumnSelectMode");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnSelectMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821992.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTextOrientation Orientation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTextOrientation>(this, "Orientation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193084.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.InlineShapes InlineShapes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.InlineShapes>(this, "InlineShapes", NetOffice.WordApi.InlineShapes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192167.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839166.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844964.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "Document", NetOffice.WordApi.Document.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836759.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ShapeRange ShapeRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShapeRange>(this, "ShapeRange", NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196937.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 NoProofing
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "NoProofing");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NoProofing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821380.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Tables TopLevelTables
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Tables>(this, "TopLevelTables", NetOffice.WordApi.Tables.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192601.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool LanguageDetected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "LanguageDetected");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "LanguageDetected", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821699.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single FitTextWidth
		{
			get
			{
				return Factory.ExecuteSinglePropertyGet(this, "FitTextWidth");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FitTextWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198226.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.HTMLDivisions HTMLDivisions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.HTMLDivisions>(this, "HTMLDivisions", NetOffice.WordApi.HTMLDivisions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.SmartTags SmartTags
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SmartTags>(this, "SmartTags", NetOffice.WordApi.SmartTags.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191940.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.ShapeRange ChildShapeRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ShapeRange>(this, "ChildShapeRange", NetOffice.WordApi.ShapeRange.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191804.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public bool HasChildShapeRange
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HasChildShapeRange");
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845098.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.FootnoteOptions FootnoteOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.FootnoteOptions>(this, "FootnoteOptions", NetOffice.WordApi.FootnoteOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192368.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.EndnoteOptions EndnoteOptions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.EndnoteOptions>(this, "EndnoteOptions", NetOffice.WordApi.EndnoteOptions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNodes XMLNodes
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNodes>(this, "XMLNodes", NetOffice.WordApi.XMLNodes.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.XMLNode XMLParentNode
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.XMLNode>(this, "XMLParentNode", NetOffice.WordApi.XMLNode.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837314.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Editors Editors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Editors>(this, "Editors", NetOffice.WordApi.Editors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838928.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public string XML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XML");
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840039.aspx </remarks>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public object EnhMetaFileBits
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "EnhMetaFileBits");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838161.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.OMaths OMaths
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.OMaths>(this, "OMaths", NetOffice.WordApi.OMaths.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820971.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public string WordOpenXML
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WordOpenXML");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ContentControls ContentControls
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControls>(this, "ContentControls", NetOffice.WordApi.ContentControls.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControl>(this, "ParentContentControl", NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845714.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192352.aspx </remarks>
		/// <param name="start">Int32 start</param>
		/// <param name="end">Int32 end</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetRange(Int32 start, Int32 end)
		{
			 Factory.ExecuteMethod(this, "SetRange", start, end);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
		/// <param name="direction">optional object direction</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Collapse(object direction)
		{
			 Factory.ExecuteMethod(this, "Collapse", direction);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834294.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Collapse()
		{
			 Factory.ExecuteMethod(this, "Collapse");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845077.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBefore(string text)
		{
			 Factory.ExecuteMethod(this, "InsertBefore", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192184.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertAfter(string text)
		{
			 Factory.ExecuteMethod(this, "InsertAfter", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next(object unit, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", NetOffice.WordApi.Range.LateBindingApiWrapperType, unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195124.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next(object unit)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", NetOffice.WordApi.Range.LateBindingApiWrapperType, unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous(object unit, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", NetOffice.WordApi.Range.LateBindingApiWrapperType, unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822303.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous(object unit)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", NetOffice.WordApi.Range.LateBindingApiWrapperType, unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf(object unit, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "StartOf", unit, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf()
		{
			return Factory.ExecuteInt32MethodGet(this, "StartOf");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196209.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "StartOf", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf(object unit, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "EndOf", unit, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf()
		{
			return Factory.ExecuteInt32MethodGet(this, "EndOf");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193383.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "EndOf", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Move(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Move()
		{
			return Factory.ExecuteInt32MethodGet(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822886.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Move(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Move", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStart", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStart");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837936.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStart", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEnd", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEnd");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845693.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEnd", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveWhile(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveWhile", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837303.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveWhile(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveWhile", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartWhile(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStartWhile", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837161.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartWhile(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStartWhile", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndWhile(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEndWhile", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837730.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndWhile(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEndWhile", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUntil(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUntil", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822578.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUntil(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUntil", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartUntil(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStartUntil", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835726.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStartUntil(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStartUntil", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
		/// <param name="cset">object cset</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndUntil(object cset, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEndUntil", cset, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839831.aspx </remarks>
		/// <param name="cset">object cset</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEndUntil(object cset)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEndUntil", cset);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192037.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196538.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840284.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak(object type)
		{
			 Factory.ExecuteMethod(this, "InsertBreak", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192797.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak()
		{
			 Factory.ExecuteMethod(this, "InsertBreak");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="link">optional object link</param>
		/// <param name="attachment">optional object attachment</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link, object attachment)
		{
			 Factory.ExecuteMethod(this, "InsertFile", new object[]{ fileName, range, confirmConversions, link, attachment });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName)
		{
			 Factory.ExecuteMethod(this, "InsertFile", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range)
		{
			 Factory.ExecuteMethod(this, "InsertFile", fileName, range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions)
		{
			 Factory.ExecuteMethod(this, "InsertFile", fileName, range, confirmConversions);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834580.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="range">optional object range</param>
		/// <param name="confirmConversions">optional object confirmConversions</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFile(string fileName, object range, object confirmConversions, object link)
		{
			 Factory.ExecuteMethod(this, "InsertFile", fileName, range, confirmConversions, link);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192633.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InStory(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "InStory", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193660.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InRange(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "InRange", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193432.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Expand", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822873.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand()
		{
			return Factory.ExecuteInt32MethodGet(this, "Expand");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837485.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraph()
		{
			 Factory.ExecuteMethod(this, "InsertParagraph");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836408.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraphAfter()
		{
			 Factory.ExecuteMethod(this, "InsertParagraphAfter");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows, numColumns);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows, numColumns, initialColumnWidth);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTableOld(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTableOld", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		/// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			 Factory.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat, insertAsField, insertAsFullWidth);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld()
		{
			 Factory.ExecuteMethod(this, "InsertDateTimeOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat)
		{
			 Factory.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTimeOld(object dateTimeFormat, object insertAsField)
		{
			 Factory.ExecuteMethod(this, "InsertDateTimeOld", dateTimeFormat, insertAsField);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional object font</param>
		/// <param name="unicode">optional object unicode</param>
		/// <param name="bias">optional object bias</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode, object bias)
		{
			 Factory.ExecuteMethod(this, "InsertSymbol", characterNumber, font, unicode, bias);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber)
		{
			 Factory.ExecuteMethod(this, "InsertSymbol", characterNumber);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional object font</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font)
		{
			 Factory.ExecuteMethod(this, "InsertSymbol", characterNumber, font);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845781.aspx </remarks>
		/// <param name="characterNumber">Int32 characterNumber</param>
		/// <param name="font">optional object font</param>
		/// <param name="unicode">optional object unicode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertSymbol(Int32 characterNumber, object font, object unicode)
		{
			 Factory.ExecuteMethod(this, "InsertSymbol", characterNumber, font, unicode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		/// <param name="includePosition">optional object includePosition</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference", new object[]{ referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		/// <param name="includePosition">optional object includePosition</param>
		/// <param name="separateNumbers">optional object separateNumbers</param>
		/// <param name="separatorString">optional object separatorString</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers, object separatorString)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference", new object[]{ referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers, separatorString });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference", referenceType, referenceKind, referenceItem);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference", referenceType, referenceKind, referenceItem, insertAsHyperlink);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193906.aspx </remarks>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		/// <param name="includePosition">optional object includePosition</param>
		/// <param name="separateNumbers">optional object separateNumbers</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCrossReference(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition, object separateNumbers)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference", new object[]{ referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition, separateNumbers });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		/// <param name="titleAutoText">optional object titleAutoText</param>
		/// <param name="position">optional object position</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText, object position)
		{
			 Factory.ExecuteMethod(this, "InsertCaption", label, title, titleAutoText, position);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		/// <param name="titleAutoText">optional object titleAutoText</param>
		/// <param name="position">optional object position</param>
		/// <param name="excludeLabel">optional object excludeLabel</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText, object position, object excludeLabel)
		{
			 Factory.ExecuteMethod(this, "InsertCaption", new object[]{ label, title, titleAutoText, position, excludeLabel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
		/// <param name="label">object label</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label)
		{
			 Factory.ExecuteMethod(this, "InsertCaption", label);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title)
		{
			 Factory.ExecuteMethod(this, "InsertCaption", label, title);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822149.aspx </remarks>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		/// <param name="titleAutoText">optional object titleAutoText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCaption(object label, object title, object titleAutoText)
		{
			 Factory.ExecuteMethod(this, "InsertCaption", label, title, titleAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840576.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CopyAsPicture()
		{
			 Factory.ExecuteMethod(this, "CopyAsPicture");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="languageID">optional object languageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object languageID)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld()
		{
			 Factory.ExecuteMethod(this, "SortOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "SortOld", excludeHeader, fieldNumber, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortOld(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "SortOld", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821863.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortAscending()
		{
			 Factory.ExecuteMethod(this, "SortAscending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845052.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortDescending()
		{
			 Factory.ExecuteMethod(this, "SortDescending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196258.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsEqual(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsEqual", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835748.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Calculate()
		{
			return Factory.ExecuteSingleMethodGet(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which, count, name);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
		/// <param name="what">optional object what</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821348.aspx </remarks>
		/// <param name="what">optional object what</param>
		/// <param name="which">optional object which</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo(object what, object which, object count)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType, what, which, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836451.aspx </remarks>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToNext", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839107.aspx </remarks>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToPrevious", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		/// <param name="placement">optional object placement</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		/// <param name="iconLabel">optional object iconLabel</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName, object iconLabel)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ iconIndex, link, placement, displayAsIcon, dataType, iconFileName, iconLabel });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial()
		{
			 Factory.ExecuteMethod(this, "PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", iconIndex);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", iconIndex, link);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		/// <param name="placement">optional object placement</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", iconIndex, link, placement);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		/// <param name="placement">optional object placement</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", iconIndex, link, placement, displayAsIcon);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		/// <param name="placement">optional object placement</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="dataType">optional object dataType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ iconIndex, link, placement, displayAsIcon, dataType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191763.aspx </remarks>
		/// <param name="iconIndex">optional object iconIndex</param>
		/// <param name="link">optional object link</param>
		/// <param name="placement">optional object placement</param>
		/// <param name="displayAsIcon">optional object displayAsIcon</param>
		/// <param name="dataType">optional object dataType</param>
		/// <param name="iconFileName">optional object iconFileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial(object iconIndex, object link, object placement, object displayAsIcon, object dataType, object iconFileName)
		{
			 Factory.ExecuteMethod(this, "PasteSpecial", new object[]{ iconIndex, link, placement, displayAsIcon, dataType, iconFileName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834516.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field PreviousField()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "PreviousField", NetOffice.WordApi.Field.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194299.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Field NextField()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Field>(this, "NextField", NetOffice.WordApi.Field.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840515.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraphBefore()
		{
			 Factory.ExecuteMethod(this, "InsertParagraphBefore");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
		/// <param name="shiftCells">optional object shiftCells</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCells(object shiftCells)
		{
			 Factory.ExecuteMethod(this, "InsertCells", shiftCells);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194778.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertCells()
		{
			 Factory.ExecuteMethod(this, "InsertCells");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
		/// <param name="character">optional object character</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Extend(object character)
		{
			 Factory.ExecuteMethod(this, "Extend", character);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821815.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Extend()
		{
			 Factory.ExecuteMethod(this, "Extend");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840081.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Shrink()
		{
			 Factory.ExecuteMethod(this, "Shrink");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit, object count, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveLeft", unit, count, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveLeft");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveLeft", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192370.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveLeft(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveLeft", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit, object count, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveRight", unit, count, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveRight");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveRight", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840899.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveRight(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveRight", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit, object count, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUp", unit, count, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUp");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUp", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194813.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveUp(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveUp", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit, object count, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveDown", unit, count, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveDown");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveDown", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838730.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="count">optional object count</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveDown(object unit, object count)
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveDown", unit, count);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey(object unit, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "HomeKey", unit, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey()
		{
			return Factory.ExecuteInt32MethodGet(this, "HomeKey");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192384.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 HomeKey(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "HomeKey", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		/// <param name="extend">optional object extend</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey(object unit, object extend)
		{
			return Factory.ExecuteInt32MethodGet(this, "EndKey", unit, extend);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey()
		{
			return Factory.ExecuteInt32MethodGet(this, "EndKey");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195593.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndKey(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "EndKey", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835736.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void EscapeKey()
		{
			 Factory.ExecuteMethod(this, "EscapeKey");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840867.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TypeText(string text)
		{
			 Factory.ExecuteMethod(this, "TypeText", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840230.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CopyFormat()
		{
			 Factory.ExecuteMethod(this, "CopyFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196637.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteFormat()
		{
			 Factory.ExecuteMethod(this, "PasteFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839799.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TypeParagraph()
		{
			 Factory.ExecuteMethod(this, "TypeParagraph");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194909.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TypeBackspace()
		{
			 Factory.ExecuteMethod(this, "TypeBackspace");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839790.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void NextSubdocument()
		{
			 Factory.ExecuteMethod(this, "NextSubdocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845750.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PreviousSubdocument()
		{
			 Factory.ExecuteMethod(this, "PreviousSubdocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836022.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectColumn()
		{
			 Factory.ExecuteMethod(this, "SelectColumn");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197469.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentFont()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentFont");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentAlignment()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentAlignment");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191872.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentSpacing()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentSpacing");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193883.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentIndent()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentIndent");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193718.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentTabs()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentTabs");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840690.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCurrentColor()
		{
			 Factory.ExecuteMethod(this, "SelectCurrentColor");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839540.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreateTextbox()
		{
			 Factory.ExecuteMethod(this, "CreateTextbox");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840046.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void WholeStory()
		{
			 Factory.ExecuteMethod(this, "WholeStory");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845469.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectRow()
		{
			 Factory.ExecuteMethod(this, "SelectRow");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196707.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SplitTable()
		{
			 Factory.ExecuteMethod(this, "SplitTable");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRows(object numRows)
		{
			 Factory.ExecuteMethod(this, "InsertRows", numRows);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193340.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRows()
		{
			 Factory.ExecuteMethod(this, "InsertRows");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838759.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertColumns()
		{
			 Factory.ExecuteMethod(this, "InsertColumns");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
		/// <param name="formula">optional object formula</param>
		/// <param name="numberFormat">optional object numberFormat</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula(object formula, object numberFormat)
		{
			 Factory.ExecuteMethod(this, "InsertFormula", formula, numberFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula()
		{
			 Factory.ExecuteMethod(this, "InsertFormula");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835475.aspx </remarks>
		/// <param name="formula">optional object formula</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertFormula(object formula)
		{
			 Factory.ExecuteMethod(this, "InsertFormula", formula);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
		/// <param name="wrap">optional object wrap</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision NextRevision(object wrap)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "NextRevision", NetOffice.WordApi.Revision.LateBindingApiWrapperType, wrap);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834850.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision NextRevision()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "NextRevision", NetOffice.WordApi.Revision.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
		/// <param name="wrap">optional object wrap</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision PreviousRevision(object wrap)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "PreviousRevision", NetOffice.WordApi.Revision.LateBindingApiWrapperType, wrap);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839603.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revision PreviousRevision()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Revision>(this, "PreviousRevision", NetOffice.WordApi.Revision.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194535.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteAsNestedTable()
		{
			 Factory.ExecuteMethod(this, "PasteAsNestedTable");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839331.aspx </remarks>
		/// <param name="name">string name</param>
		/// <param name="styleName">string styleName</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.AutoTextEntry CreateAutoTextEntry(string name, string styleName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.AutoTextEntry>(this, "CreateAutoTextEntry", NetOffice.WordApi.AutoTextEntry.LateBindingApiWrapperType, name, styleName);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838494.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DetectLanguage()
		{
			 Factory.ExecuteMethod(this, "DetectLanguage");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195143.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SelectCell()
		{
			 Factory.ExecuteMethod(this, "SelectCell");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsBelow(object numRows)
		{
			 Factory.ExecuteMethod(this, "InsertRowsBelow", numRows);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838718.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsBelow()
		{
			 Factory.ExecuteMethod(this, "InsertRowsBelow");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844950.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertColumnsRight()
		{
			 Factory.ExecuteMethod(this, "InsertColumnsRight");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
		/// <param name="numRows">optional object numRows</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsAbove(object numRows)
		{
			 Factory.ExecuteMethod(this, "InsertRowsAbove", numRows);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840557.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertRowsAbove()
		{
			 Factory.ExecuteMethod(this, "InsertRowsAbove");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821034.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RtlRun()
		{
			 Factory.ExecuteMethod(this, "RtlRun");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839502.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void LtrRun()
		{
			 Factory.ExecuteMethod(this, "LtrRun");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845275.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void BoldRun()
		{
			 Factory.ExecuteMethod(this, "BoldRun");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845442.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ItalicRun()
		{
			 Factory.ExecuteMethod(this, "ItalicRun");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836904.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void RtlPara()
		{
			 Factory.ExecuteMethod(this, "RtlPara");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834853.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void LtrPara()
		{
			 Factory.ExecuteMethod(this, "LtrPara");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		/// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
		/// <param name="dateLanguage">optional object dateLanguage</param>
		/// <param name="calendarType">optional object calendarType</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage, object calendarType)
		{
			 Factory.ExecuteMethod(this, "InsertDateTime", new object[]{ dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage, calendarType });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime()
		{
			 Factory.ExecuteMethod(this, "InsertDateTime");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat)
		{
			 Factory.ExecuteMethod(this, "InsertDateTime", dateTimeFormat);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField)
		{
			 Factory.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		/// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth)
		{
			 Factory.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField, insertAsFullWidth);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840789.aspx </remarks>
		/// <param name="dateTimeFormat">optional object dateTimeFormat</param>
		/// <param name="insertAsField">optional object insertAsField</param>
		/// <param name="insertAsFullWidth">optional object insertAsFullWidth</param>
		/// <param name="dateLanguage">optional object dateLanguage</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime(object dateTimeFormat, object insertAsField, object insertAsFullWidth, object dateLanguage)
		{
			 Factory.ExecuteMethod(this, "InsertDateTime", dateTimeFormat, insertAsField, insertAsFullWidth, dateLanguage);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		/// <param name="subFieldNumber">optional object subFieldNumber</param>
		/// <param name="subFieldNumber2">optional object subFieldNumber2</param>
		/// <param name="subFieldNumber3">optional object subFieldNumber3</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2, object subFieldNumber3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2, subFieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			 Factory.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "Sort", excludeHeader, fieldNumber, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		/// <param name="subFieldNumber">optional object subFieldNumber</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194250.aspx </remarks>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		/// <param name="subFieldNumber">optional object subFieldNumber</param>
		/// <param name="subFieldNumber2">optional object subFieldNumber2</param>
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID, object subFieldNumber, object subFieldNumber2)
		{
			 Factory.ExecuteMethod(this, "Sort", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID, subFieldNumber, subFieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		/// <param name="autoFitBehavior">optional object autoFitBehavior</param>
		/// <param name="defaultTableBehavior">optional object defaultTableBehavior</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior, object defaultTableBehavior)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior, defaultTableBehavior });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows, numColumns);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, separator, numRows, numColumns, initialColumnWidth);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836746.aspx </remarks>
		/// <param name="separator">optional object separator</param>
		/// <param name="numRows">optional object numRows</param>
		/// <param name="numColumns">optional object numColumns</param>
		/// <param name="initialColumnWidth">optional object initialColumnWidth</param>
		/// <param name="format">optional object format</param>
		/// <param name="applyBorders">optional object applyBorders</param>
		/// <param name="applyShading">optional object applyShading</param>
		/// <param name="applyFont">optional object applyFont</param>
		/// <param name="applyColor">optional object applyColor</param>
		/// <param name="applyHeadingRows">optional object applyHeadingRows</param>
		/// <param name="applyLastRow">optional object applyLastRow</param>
		/// <param name="applyFirstColumn">optional object applyFirstColumn</param>
		/// <param name="applyLastColumn">optional object applyLastColumn</param>
		/// <param name="autoFit">optional object autoFit</param>
		/// <param name="autoFitBehavior">optional object autoFitBehavior</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable(object separator, object numRows, object numColumns, object initialColumnWidth, object format, object applyBorders, object applyShading, object applyFont, object applyColor, object applyHeadingRows, object applyLastRow, object applyFirstColumn, object applyLastColumn, object autoFit, object autoFitBehavior)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType, new object[]{ separator, numRows, numColumns, initialColumnWidth, format, applyBorders, applyShading, applyFont, applyColor, applyHeadingRows, applyLastRow, applyFirstColumn, applyLastColumn, autoFit, autoFitBehavior });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000()
		{
			 Factory.ExecuteMethod(this, "Sort2000");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader)
		{
			 Factory.ExecuteMethod(this, "Sort2000", excludeHeader);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber)
		{
			 Factory.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber, sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "Sort2000", excludeHeader, fieldNumber, sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3 });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="excludeHeader">optional object excludeHeader</param>
		/// <param name="fieldNumber">optional object fieldNumber</param>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="fieldNumber2">optional object fieldNumber2</param>
		/// <param name="sortFieldType2">optional object sortFieldType2</param>
		/// <param name="sortOrder2">optional object sortOrder2</param>
		/// <param name="fieldNumber3">optional object fieldNumber3</param>
		/// <param name="sortFieldType3">optional object sortFieldType3</param>
		/// <param name="sortOrder3">optional object sortOrder3</param>
		/// <param name="sortColumn">optional object sortColumn</param>
		/// <param name="separator">optional object separator</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void Sort2000(object excludeHeader, object fieldNumber, object sortFieldType, object sortOrder, object fieldNumber2, object sortFieldType2, object sortOrder2, object fieldNumber3, object sortFieldType3, object sortOrder3, object sortColumn, object separator, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 Factory.ExecuteMethod(this, "Sort2000", new object[]{ excludeHeader, fieldNumber, sortFieldType, sortOrder, fieldNumber2, sortFieldType2, sortOrder2, fieldNumber3, sortFieldType3, sortOrder3, sortColumn, separator, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197496.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ClearFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196969.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PasteAppendTable()
		{
			 Factory.ExecuteMethod(this, "PasteAppendTable");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839633.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ToggleCharacterCode()
		{
			 Factory.ExecuteMethod(this, "ToggleCharacterCode");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821674.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
		{
			 Factory.ExecuteMethod(this, "PasteAndFormat", type);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837670.aspx </remarks>
		/// <param name="linkedToExcel">bool linkedToExcel</param>
		/// <param name="wordFormatting">bool wordFormatting</param>
		/// <param name="rTF">bool rTF</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PasteExcelTable(bool linkedToExcel, bool wordFormatting, bool rTF)
		{
			 Factory.ExecuteMethod(this, "PasteExcelTable", linkedToExcel, wordFormatting, rTF);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838352.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void ShrinkDiscontiguousSelection()
		{
			 Factory.ExecuteMethod(this, "ShrinkDiscontiguousSelection");
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838293.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void InsertStyleSeparator()
		{
			 Factory.ExecuteMethod(this, "InsertStyleSeparator");
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		/// <param name="includePosition">optional object includePosition</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink, object includePosition)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference_2002", new object[]{ referenceType, referenceKind, referenceItem, insertAsHyperlink, includePosition });
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference_2002", referenceType, referenceKind, referenceItem);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="referenceType">object referenceType</param>
		/// <param name="referenceKind">NetOffice.WordApi.Enums.WdReferenceKind referenceKind</param>
		/// <param name="referenceItem">object referenceItem</param>
		/// <param name="insertAsHyperlink">optional object insertAsHyperlink</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCrossReference_2002(object referenceType, NetOffice.WordApi.Enums.WdReferenceKind referenceKind, object referenceItem, object insertAsHyperlink)
		{
			 Factory.ExecuteMethod(this, "InsertCrossReference_2002", referenceType, referenceKind, referenceItem, insertAsHyperlink);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		/// <param name="titleAutoText">optional object titleAutoText</param>
		/// <param name="position">optional object position</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title, object titleAutoText, object position)
		{
			 Factory.ExecuteMethod(this, "InsertCaptionXP", label, title, titleAutoText, position);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="label">object label</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label)
		{
			 Factory.ExecuteMethod(this, "InsertCaptionXP", label);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title)
		{
			 Factory.ExecuteMethod(this, "InsertCaptionXP", label, title);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="label">object label</param>
		/// <param name="title">optional object title</param>
		/// <param name="titleAutoText">optional object titleAutoText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertCaptionXP(object label, object title, object titleAutoText)
		{
			 Factory.ExecuteMethod(this, "InsertCaptionXP", label, title, titleAutoText);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange(object editorID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", NetOffice.WordApi.Range.LateBindingApiWrapperType, editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844866.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
		/// <param name="xML">string xML</param>
		/// <param name="transform">optional object transform</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertXML(string xML, object transform)
		{
			 Factory.ExecuteMethod(this, "InsertXML", xML, transform);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821369.aspx </remarks>
		/// <param name="xML">string xML</param>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public void InsertXML(string xML)
		{
			 Factory.ExecuteMethod(this, "InsertXML", xML);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838493.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearParagraphStyle()
		{
			 Factory.ExecuteMethod(this, "ClearParagraphStyle");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191975.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearCharacterAllFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearCharacterAllFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841083.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearCharacterStyle()
		{
			 Factory.ExecuteMethod(this, "ClearCharacterStyle");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838672.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearCharacterDirectFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearCharacterDirectFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		/// <param name="fixedFormatExtClassPtr">optional object fixedFormatExtClassPtr</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1, object fixedFormatExtClassPtr)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1, fixedFormatExtClassPtr });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", outputFileName, exportFormat, openAfterExport, optimizeFor);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845579.aspx </remarks>
		/// <param name="outputFileName">string outputFileName</param>
		/// <param name="exportFormat">NetOffice.WordApi.Enums.WdExportFormat exportFormat</param>
		/// <param name="openAfterExport">optional bool OpenAfterExport = false</param>
		/// <param name="optimizeFor">optional NetOffice.WordApi.Enums.WdExportOptimizeFor OptimizeFor = 0</param>
		/// <param name="exportCurrentPage">optional bool ExportCurrentPage = false</param>
		/// <param name="item">optional NetOffice.WordApi.Enums.WdExportItem Item = 0</param>
		/// <param name="includeDocProps">optional bool IncludeDocProps = false</param>
		/// <param name="keepIRM">optional bool KeepIRM = true</param>
		/// <param name="createBookmarks">optional NetOffice.WordApi.Enums.WdExportCreateBookmarks CreateBookmarks = 0</param>
		/// <param name="docStructureTags">optional bool DocStructureTags = true</param>
		/// <param name="bitmapMissingFonts">optional bool BitmapMissingFonts = true</param>
		/// <param name="useISO19005_1">optional bool UseISO19005_1 = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportAsFixedFormat(string outputFileName, NetOffice.WordApi.Enums.WdExportFormat exportFormat, object openAfterExport, object optimizeFor, object exportCurrentPage, object item, object includeDocProps, object keepIRM, object createBookmarks, object docStructureTags, object bitmapMissingFonts, object useISO19005_1)
		{
			 Factory.ExecuteMethod(this, "ExportAsFixedFormat", new object[]{ outputFileName, exportFormat, openAfterExport, optimizeFor, exportCurrentPage, item, includeDocProps, keepIRM, createBookmarks, docStructureTags, bitmapMissingFonts, useISO19005_1 });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196419.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ReadingModeGrowFont()
		{
			 Factory.ExecuteMethod(this, "ReadingModeGrowFont");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196279.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ReadingModeShrinkFont()
		{
			 Factory.ExecuteMethod(this, "ReadingModeShrinkFont");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836876.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearParagraphAllFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearParagraphAllFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197502.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ClearParagraphDirectFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearParagraphDirectFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195985.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public void InsertNewPage()
		{
			 Factory.ExecuteMethod(this, "InsertNewPage");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		/// <param name="languageID">optional object languageID</param>
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe, object languageID)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", new object[]{ sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe, languageID });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings()
		{
			 Factory.ExecuteMethod(this, "SortByHeadings");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", sortFieldType);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder, caseSensitive);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", sortFieldType, sortOrder, caseSensitive, bidiSort);
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", new object[]{ sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", new object[]{ sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", new object[]{ sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232374.aspx </remarks>
		/// <param name="sortFieldType">optional object sortFieldType</param>
		/// <param name="sortOrder">optional object sortOrder</param>
		/// <param name="caseSensitive">optional object caseSensitive</param>
		/// <param name="bidiSort">optional object bidiSort</param>
		/// <param name="ignoreThe">optional object ignoreThe</param>
		/// <param name="ignoreKashida">optional object ignoreKashida</param>
		/// <param name="ignoreDiacritics">optional object ignoreDiacritics</param>
		/// <param name="ignoreHe">optional object ignoreHe</param>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings(object sortFieldType, object sortOrder, object caseSensitive, object bidiSort, object ignoreThe, object ignoreKashida, object ignoreDiacritics, object ignoreHe)
		{
			 Factory.ExecuteMethod(this, "SortByHeadings", new object[]{ sortFieldType, sortOrder, caseSensitive, bidiSort, ignoreThe, ignoreKashida, ignoreDiacritics, ignoreHe });
		}

		#endregion

		#pragma warning restore
	}
}
