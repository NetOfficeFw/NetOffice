using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// Range
	/// </summary>
	[SyntaxBypass]
 	public class Range_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Range_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Range_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="dataOnly">optional bool dataOnly</param>
		/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
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
	/// DispatchInterface Range 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845882.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Range : Range_
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
                    _type = typeof(Range);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Range(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Range(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Range(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195101.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192541.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836102.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840998.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837543.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Duplicate
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "Duplicate", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837652.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191956.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836346.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840991.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845462.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196597.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193114.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192150.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836072.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834837.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837006.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835448.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839529.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.TextRetrievalMode TextRetrievalMode
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.TextRetrievalMode>(this, "TextRetrievalMode", NetOffice.WordApi.TextRetrievalMode.LateBindingApiWrapperType);
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "TextRetrievalMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845620.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834816.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837877.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834843.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195640.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListFormat ListFormat
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListFormat>(this, "ListFormat", NetOffice.WordApi.ListFormat.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195181.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196242.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839336.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844991.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Bold
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Bold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Bold", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821583.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Italic
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Italic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Italic", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821959.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdUnderline Underline
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdUnderline>(this, "Underline");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Underline", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198151.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdEmphasisMark EmphasisMark
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdEmphasisMark>(this, "EmphasisMark");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EmphasisMark", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844978.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool DisableCharacterSpaceGrid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisableCharacterSpaceGrid");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisableCharacterSpaceGrid", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838481.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Revisions Revisions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Revisions>(this, "Revisions", NetOffice.WordApi.Revisions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836418.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845486.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839161.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837028.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SynonymInfo SynonymInfo
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.SynonymInfo>(this, "SynonymInfo", NetOffice.WordApi.SynonymInfo.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838128.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838758.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ListParagraphs ListParagraphs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ListParagraphs>(this, "ListParagraphs", NetOffice.WordApi.ListParagraphs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837692.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Subdocuments Subdocuments
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Subdocuments>(this, "Subdocuments", NetOffice.WordApi.Subdocuments.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840317.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool GrammarChecked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GrammarChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GrammarChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196502.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool SpellingChecked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SpellingChecked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SpellingChecked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841064.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdColorIndex HighlightColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdColorIndex>(this, "HighlightColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HighlightColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197474.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840908.aspx </remarks>
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
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CanEdit
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CanEdit");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CanPaste
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CanPaste");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845343.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845646.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191844.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195912.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192629.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837242.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834838.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCharacterCase Case
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCharacterCase>(this, "Case");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Case", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834587.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdInformation type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), Redirect("get_Information")]
		public object Information(NetOffice.WordApi.Enums.WdInformation type)
		{
			return get_Information(type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837707.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ReadabilityStatistics ReadabilityStatistics
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ReadabilityStatistics>(this, "ReadabilityStatistics", NetOffice.WordApi.ReadabilityStatistics.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192406.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors GrammaticalErrors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "GrammaticalErrors", NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195285.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.ProofreadingErrors SpellingErrors
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ProofreadingErrors>(this, "SpellingErrors", NetOffice.WordApi.ProofreadingErrors.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195776.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195321.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193730.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range NextStoryRange
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Range>(this, "NextStoryRange", NetOffice.WordApi.Range.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193321.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844803.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839349.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197181.aspx </remarks>
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdHorizontalInVerticalType HorizontalInVertical
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdHorizontalInVerticalType>(this, "HorizontalInVertical");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HorizontalInVertical", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845231.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdTwoLinesInOneType TwoLinesInOne
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdTwoLinesInOneType>(this, "TwoLinesInOne");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TwoLinesInOne", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195015.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CombineCharacters
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CombineCharacters");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CombineCharacters", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844920.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194640.aspx </remarks>
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
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192353.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Scripts Scripts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Scripts>(this, "Scripts", NetOffice.OfficeApi.Scripts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822135.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdCharacterWidth CharacterWidth
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdCharacterWidth>(this, "CharacterWidth");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "CharacterWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840112.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdKana Kana
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdKana>(this, "Kana");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Kana", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821869.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 BoldBi
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BoldBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BoldBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197717.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ItalicBi
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ItalicBi");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ItalicBi", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196542.aspx </remarks>
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194856.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820977.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
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
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194311.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public NetOffice.WordApi.Document Document
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Document>(this, "Document", NetOffice.WordApi.Document.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195199.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195039.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840972.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193656.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192034.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822393.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192339.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public object CharacterStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "CharacterStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196075.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public object ParagraphStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ParagraphStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196585.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public object ListStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ListStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841045.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public object TableStyle
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TableStyle");
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839822.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837448.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839629.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public NetOffice.WordApi.ContentControl ParentContentControl
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ContentControl>(this, "ParentContentControl", NetOffice.WordApi.ContentControl.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845600.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.CoAuthLocks Locks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthLocks>(this, "Locks", NetOffice.WordApi.CoAuthLocks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196284.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.CoAuthUpdates Updates
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.CoAuthUpdates>(this, "Updates", NetOffice.WordApi.CoAuthUpdates.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823246.aspx </remarks>
		[SupportByVersion("Word", 14,15,16)]
		public NetOffice.WordApi.Conflicts Conflicts
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Conflicts>(this, "Conflicts", NetOffice.WordApi.Conflicts.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231893.aspx </remarks>
		[SupportByVersion("Word", 15, 16)]
		public Int32 TextVisibleOnScreen
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextVisibleOnScreen");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820813.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Select()
		{
			 Factory.ExecuteMethod(this, "Select");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823262.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
		/// <param name="direction">optional object direction</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Collapse(object direction)
		{
			 Factory.ExecuteMethod(this, "Collapse", direction);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840825.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Collapse()
		{
			 Factory.ExecuteMethod(this, "Collapse");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836272.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBefore(string text)
		{
			 Factory.ExecuteMethod(this, "InsertBefore", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192427.aspx </remarks>
		/// <param name="text">string text</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertAfter(string text)
		{
			 Factory.ExecuteMethod(this, "InsertAfter", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Next()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Next", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822953.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range Previous()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "Previous", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840143.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 StartOf()
		{
			return Factory.ExecuteInt32MethodGet(this, "StartOf");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195382.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 EndOf()
		{
			return Factory.ExecuteInt32MethodGet(this, "EndOf");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837285.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Move()
		{
			return Factory.ExecuteInt32MethodGet(this, "Move");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194352.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveStart()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveStart");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823249.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 MoveEnd()
		{
			return Factory.ExecuteInt32MethodGet(this, "MoveEnd");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194698.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192586.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838537.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835396.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840312.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192403.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197156.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195686.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837718.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845105.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
		/// <param name="type">optional object type</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak(object type)
		{
			 Factory.ExecuteMethod(this, "InsertBreak", type);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835132.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertBreak()
		{
			 Factory.ExecuteMethod(this, "InsertBreak");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835231.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197125.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InStory(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "InStory", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822960.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool InRange(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "InRange", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Delete()
		{
			return Factory.ExecuteInt32MethodGet(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845114.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837449.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void WholeStory()
		{
			 Factory.ExecuteMethod(this, "WholeStory");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
		/// <param name="unit">optional object unit</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand(object unit)
		{
			return Factory.ExecuteInt32MethodGet(this, "Expand", unit);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838477.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Expand()
		{
			return Factory.ExecuteInt32MethodGet(this, "Expand");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196197.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraph()
		{
			 Factory.ExecuteMethod(this, "InsertParagraph");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822546.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193081.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196302.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841026.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836633.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193013.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortAscending()
		{
			 Factory.ExecuteMethod(this, "SortAscending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844858.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SortDescending()
		{
			 Factory.ExecuteMethod(this, "SortDescending");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838323.aspx </remarks>
		/// <param name="range">NetOffice.WordApi.Range range</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool IsEqual(NetOffice.WordApi.Range range)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsEqual", range);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821015.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Single Calculate()
		{
			return Factory.ExecuteSingleMethodGet(this, "Calculate");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoTo()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoTo", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835184.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844826.aspx </remarks>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToNext(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToNext", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836673.aspx </remarks>
		/// <param name="what">NetOffice.WordApi.Enums.WdGoToItem what</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToPrevious(NetOffice.WordApi.Enums.WdGoToItem what)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToPrevious", NetOffice.WordApi.Range.LateBindingApiWrapperType, what);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteSpecial()
		{
			 Factory.ExecuteMethod(this, "PasteSpecial");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821124.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835691.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void LookupNameProperties()
		{
			 Factory.ExecuteMethod(this, "LookupNameProperties");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196924.aspx </remarks>
		/// <param name="statistic">NetOffice.WordApi.Enums.WdStatistic statistic</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 ComputeStatistics(NetOffice.WordApi.Enums.WdStatistic statistic)
		{
			return Factory.ExecuteInt32MethodGet(this, "ComputeStatistics", statistic);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192827.aspx </remarks>
		/// <param name="direction">Int32 direction</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Relocate(Int32 direction)
		{
			 Factory.ExecuteMethod(this, "Relocate", direction);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839497.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSynonyms()
		{
			 Factory.ExecuteMethod(this, "CheckSynonyms");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">string edition</param>
		/// <param name="format">optional object format</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SubscribeTo(string edition, object format)
		{
			 Factory.ExecuteMethod(this, "SubscribeTo", edition, format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">string edition</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SubscribeTo(string edition)
		{
			 Factory.ExecuteMethod(this, "SubscribeTo", edition);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		/// <param name="containsRTF">optional object containsRTF</param>
		/// <param name="containsText">optional object containsText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object containsPICT, object containsRTF, object containsText)
		{
			 Factory.ExecuteMethod(this, "CreatePublisher", edition, containsPICT, containsRTF, containsText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreatePublisher()
		{
			 Factory.ExecuteMethod(this, "CreatePublisher");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition)
		{
			 Factory.ExecuteMethod(this, "CreatePublisher", edition);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object containsPICT)
		{
			 Factory.ExecuteMethod(this, "CreatePublisher", edition, containsPICT);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="edition">optional object edition</param>
		/// <param name="containsPICT">optional object containsPICT</param>
		/// <param name="containsRTF">optional object containsRTF</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CreatePublisher(object edition, object containsPICT, object containsRTF)
		{
			 Factory.ExecuteMethod(this, "CreatePublisher", edition, containsPICT, containsRTF);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838952.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertAutoText()
		{
			 Factory.ExecuteMethod(this, "InsertAutoText");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="dataSource">optional object dataSource</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		/// <param name="includeFields">optional object includeFields</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to, object includeFields)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to, includeFields });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase()
		{
			 Factory.ExecuteMethod(this, "InsertDatabase");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", format);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", format, style);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", format, style, linkToSource);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", format, style, linkToSource, connection);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="dataSource">optional object dataSource</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="dataSource">optional object dataSource</param>
		/// <param name="from">optional object from</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838122.aspx </remarks>
		/// <param name="format">optional object format</param>
		/// <param name="style">optional object style</param>
		/// <param name="linkToSource">optional object linkToSource</param>
		/// <param name="connection">optional object connection</param>
		/// <param name="sQLStatement">optional object sQLStatement</param>
		/// <param name="sQLStatement1">optional object sQLStatement1</param>
		/// <param name="passwordDocument">optional object passwordDocument</param>
		/// <param name="passwordTemplate">optional object passwordTemplate</param>
		/// <param name="writePasswordDocument">optional object writePasswordDocument</param>
		/// <param name="writePasswordTemplate">optional object writePasswordTemplate</param>
		/// <param name="dataSource">optional object dataSource</param>
		/// <param name="from">optional object from</param>
		/// <param name="to">optional object to</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDatabase(object format, object style, object linkToSource, object connection, object sQLStatement, object sQLStatement1, object passwordDocument, object passwordTemplate, object writePasswordDocument, object writePasswordTemplate, object dataSource, object from, object to)
		{
			 Factory.ExecuteMethod(this, "InsertDatabase", new object[]{ format, style, linkToSource, connection, sQLStatement, sQLStatement1, passwordDocument, passwordTemplate, writePasswordDocument, writePasswordTemplate, dataSource, from, to });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845283.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void AutoFormat()
		{
			 Factory.ExecuteMethod(this, "AutoFormat");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193931.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckGrammar()
		{
			 Factory.ExecuteMethod(this, "CheckGrammar");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		/// <param name="customDictionary10">optional object customDictionary10</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling()
		{
			 Factory.ExecuteMethod(this, "CheckSpelling");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194400.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="alwaysSuggest">optional object alwaysSuggest</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void CheckSpelling(object customDictionary, object ignoreUppercase, object alwaysSuggest, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			 Factory.ExecuteMethod(this, "CheckSpelling", new object[]{ customDictionary, ignoreUppercase, alwaysSuggest, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		/// <param name="customDictionary10">optional object customDictionary10</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9, object customDictionary10)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9, customDictionary10 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, customDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, customDictionary, ignoreUppercase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, customDictionary, ignoreUppercase, mainDictionary);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, customDictionary, ignoreUppercase, mainDictionary, suggestionMode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196525.aspx </remarks>
		/// <param name="customDictionary">optional object customDictionary</param>
		/// <param name="ignoreUppercase">optional object ignoreUppercase</param>
		/// <param name="mainDictionary">optional object mainDictionary</param>
		/// <param name="suggestionMode">optional object suggestionMode</param>
		/// <param name="customDictionary2">optional object customDictionary2</param>
		/// <param name="customDictionary3">optional object customDictionary3</param>
		/// <param name="customDictionary4">optional object customDictionary4</param>
		/// <param name="customDictionary5">optional object customDictionary5</param>
		/// <param name="customDictionary6">optional object customDictionary6</param>
		/// <param name="customDictionary7">optional object customDictionary7</param>
		/// <param name="customDictionary8">optional object customDictionary8</param>
		/// <param name="customDictionary9">optional object customDictionary9</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.SpellingSuggestions GetSpellingSuggestions(object customDictionary, object ignoreUppercase, object mainDictionary, object suggestionMode, object customDictionary2, object customDictionary3, object customDictionary4, object customDictionary5, object customDictionary6, object customDictionary7, object customDictionary8, object customDictionary9)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.SpellingSuggestions>(this, "GetSpellingSuggestions", NetOffice.WordApi.SpellingSuggestions.LateBindingApiWrapperType, new object[]{ customDictionary, ignoreUppercase, mainDictionary, suggestionMode, customDictionary2, customDictionary3, customDictionary4, customDictionary5, customDictionary6, customDictionary7, customDictionary8, customDictionary9 });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821256.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertParagraphBefore()
		{
			 Factory.ExecuteMethod(this, "InsertParagraphBefore");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195326.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void NextSubdocument()
		{
			 Factory.ExecuteMethod(this, "NextSubdocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195945.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PreviousSubdocument()
		{
			 Factory.ExecuteMethod(this, "PreviousSubdocument");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		/// <param name="conversionsMode">optional object conversionsMode</param>
		/// <param name="fastConversion">optional object fastConversion</param>
		/// <param name="checkHangulEnding">optional object checkHangulEnding</param>
		/// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
		/// <param name="customDictionary">optional object customDictionary</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering, object customDictionary)
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja", new object[]{ conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering, customDictionary });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja()
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		/// <param name="conversionsMode">optional object conversionsMode</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja(object conversionsMode)
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		/// <param name="conversionsMode">optional object conversionsMode</param>
		/// <param name="fastConversion">optional object fastConversion</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion)
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		/// <param name="conversionsMode">optional object conversionsMode</param>
		/// <param name="fastConversion">optional object fastConversion</param>
		/// <param name="checkHangulEnding">optional object checkHangulEnding</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding)
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion, checkHangulEnding);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192769.aspx </remarks>
		/// <param name="conversionsMode">optional object conversionsMode</param>
		/// <param name="fastConversion">optional object fastConversion</param>
		/// <param name="checkHangulEnding">optional object checkHangulEnding</param>
		/// <param name="enableRecentOrdering">optional object enableRecentOrdering</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ConvertHangulAndHanja(object conversionsMode, object fastConversion, object checkHangulEnding, object enableRecentOrdering)
		{
			 Factory.ExecuteMethod(this, "ConvertHangulAndHanja", conversionsMode, fastConversion, checkHangulEnding, enableRecentOrdering);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822962.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PasteAsNestedTable()
		{
			 Factory.ExecuteMethod(this, "PasteAsNestedTable");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
		/// <param name="style">object style</param>
		/// <param name="symbol">optional object symbol</param>
		/// <param name="enclosedText">optional object enclosedText</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ModifyEnclosure(object style, object symbol, object enclosedText)
		{
			 Factory.ExecuteMethod(this, "ModifyEnclosure", style, symbol, enclosedText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
		/// <param name="style">object style</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ModifyEnclosure(object style)
		{
			 Factory.ExecuteMethod(this, "ModifyEnclosure", style);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191734.aspx </remarks>
		/// <param name="style">object style</param>
		/// <param name="symbol">optional object symbol</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ModifyEnclosure(object style, object symbol)
		{
			 Factory.ExecuteMethod(this, "ModifyEnclosure", style, symbol);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		/// <param name="fontSize">optional Int32 FontSize = 0</param>
		/// <param name="fontName">optional string FontName = </param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PhoneticGuide(string text, object alignment, object raise, object fontSize, object fontName)
		{
			 Factory.ExecuteMethod(this, "PhoneticGuide", new object[]{ text, alignment, raise, fontSize, fontName });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
		/// <param name="text">string text</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PhoneticGuide(string text)
		{
			 Factory.ExecuteMethod(this, "PhoneticGuide", text);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PhoneticGuide(string text, object alignment)
		{
			 Factory.ExecuteMethod(this, "PhoneticGuide", text, alignment);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PhoneticGuide(string text, object alignment, object raise)
		{
			 Factory.ExecuteMethod(this, "PhoneticGuide", text, alignment, raise);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840645.aspx </remarks>
		/// <param name="text">string text</param>
		/// <param name="alignment">optional NetOffice.WordApi.Enums.WdPhoneticGuideAlignmentType Alignment = -1</param>
		/// <param name="raise">optional Int32 Raise = 0</param>
		/// <param name="fontSize">optional Int32 FontSize = 0</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void PhoneticGuide(string text, object alignment, object raise, object fontSize)
		{
			 Factory.ExecuteMethod(this, "PhoneticGuide", text, alignment, raise, fontSize);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void InsertDateTime()
		{
			 Factory.ExecuteMethod(this, "InsertDateTime");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192209.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void Sort()
		{
			 Factory.ExecuteMethod(this, "Sort");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192159.aspx </remarks>
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195289.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void DetectLanguage()
		{
			 Factory.ExecuteMethod(this, "DetectLanguage");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Table ConvertToTable()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Table>(this, "ConvertToTable", NetOffice.WordApi.Table.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835980.aspx </remarks>
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
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		/// <param name="commonTerms">optional bool CommonTerms = false</param>
		/// <param name="useVariants">optional bool UseVariants = false</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TCSCConverter(object wdTCSCConverterDirection, object commonTerms, object useVariants)
		{
			 Factory.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection, commonTerms, useVariants);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TCSCConverter()
		{
			 Factory.ExecuteMethod(this, "TCSCConverter");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TCSCConverter(object wdTCSCConverterDirection)
		{
			 Factory.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198332.aspx </remarks>
		/// <param name="wdTCSCConverterDirection">optional NetOffice.WordApi.Enums.WdTCSCConverterDirection WdTCSCConverterDirection = 2</param>
		/// <param name="commonTerms">optional bool CommonTerms = false</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void TCSCConverter(object wdTCSCConverterDirection, object commonTerms)
		{
			 Factory.ExecuteMethod(this, "TCSCConverter", wdTCSCConverterDirection, commonTerms);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193749.aspx </remarks>
		/// <param name="type">NetOffice.WordApi.Enums.WdRecoveryType type</param>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PasteAndFormat(NetOffice.WordApi.Enums.WdRecoveryType type)
		{
			 Factory.ExecuteMethod(this, "PasteAndFormat", type);
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193063.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839173.aspx </remarks>
		[SupportByVersion("Word", 10,11,12,14,15,16)]
		public void PasteAppendTable()
		{
			 Factory.ExecuteMethod(this, "PasteAppendTable");
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
		/// <param name="editorID">optional object editorID</param>
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange(object editorID)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", NetOffice.WordApi.Range.LateBindingApiWrapperType, editorID);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195065.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 11,12,14,15,16)]
		public NetOffice.WordApi.Range GoToEditableRange()
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.WordApi.Range>(this, "GoToEditableRange", NetOffice.WordApi.Range.LateBindingApiWrapperType);
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839129.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822335.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="format">NetOffice.WordApi.Enums.WdSaveFormat format</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ExportFragment(string fileName, NetOffice.WordApi.Enums.WdSaveFormat format)
		{
			 Factory.ExecuteMethod(this, "ExportFragment", fileName, format);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821878.aspx </remarks>
		/// <param name="level">Int16 level</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void SetListLevel(Int16 level)
		{
			 Factory.ExecuteMethod(this, "SetListLevel", level);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
		/// <param name="alignment">Int32 alignment</param>
		/// <param name="relativeTo">optional Int32 RelativeTo = 0</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void InsertAlignmentTab(Int32 alignment, object relativeTo)
		{
			 Factory.ExecuteMethod(this, "InsertAlignmentTab", alignment, relativeTo);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191966.aspx </remarks>
		/// <param name="alignment">Int32 alignment</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void InsertAlignmentTab(Int32 alignment)
		{
			 Factory.ExecuteMethod(this, "InsertAlignmentTab", alignment);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		/// <param name="matchDestination">optional bool MatchDestination = false</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public void ImportFragment(string fileName, object matchDestination)
		{
			 Factory.ExecuteMethod(this, "ImportFragment", fileName, matchDestination);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839096.aspx </remarks>
		/// <param name="fileName">string fileName</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public void ImportFragment(string fileName)
		{
			 Factory.ExecuteMethod(this, "ImportFragment", fileName);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838566.aspx </remarks>
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
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 15, 16)]
		public void SortByHeadings()
		{
			 Factory.ExecuteMethod(this, "SortByHeadings");
		}

		/// <summary>
		/// SupportByVersion Word 15,16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230923.aspx </remarks>
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
