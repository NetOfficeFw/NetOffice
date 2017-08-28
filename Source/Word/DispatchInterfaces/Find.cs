using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi
{
	/// <summary>
	/// DispatchInterface Find 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839118.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Find : COMObject
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
                    _type = typeof(Find);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Find(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Find(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Find(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196396.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839624.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834556.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839325.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Forward
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Forward");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Forward", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822678.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838143.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Found
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Found");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845697.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchAllWordForms
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchAllWordForms");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchAllWordForms", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837923.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchCase
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchCase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchCase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838695.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchWildcards
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchWildcards");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchWildcards", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821942.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchSoundsLike
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchSoundsLike");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchSoundsLike", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835745.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchWholeWord
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchWholeWord");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchWholeWord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821682.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchFuzzy
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFuzzy");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchFuzzy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838094.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchByte
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchByte");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchByte", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836406.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192137.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838976.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837887.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821028.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public Int32 Highlight
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Highlight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Highlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836618.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Replacement Replacement
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Replacement>(this, "Replacement", NetOffice.WordApi.Replacement.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197498.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Frame Frame
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frame>(this, "Frame", NetOffice.WordApi.Frame.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public NetOffice.WordApi.Enums.WdFindWrap Wrap
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFindWrap>(this, "Wrap");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Wrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834863.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Format
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Format");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195137.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836860.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821910.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool CorrectHangulEndings
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CorrectHangulEndings");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CorrectHangulEndings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195417.aspx </remarks>
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
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845200.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchKashida
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchKashida");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchKashida", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839133.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchDiacritics
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchDiacritics");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchDiacritics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845597.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchAlefHamza
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchAlefHamza");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchAlefHamza", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool MatchControl
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchControl");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191768.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool MatchPhrase
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchPhrase");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchPhrase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197820.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool MatchPrefix
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchPrefix");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchPrefix", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839710.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool MatchSuffix
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchSuffix");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchSuffix", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821316.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool IgnoreSpace
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreSpace");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194518.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool IgnorePunct
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnorePunct");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnorePunct", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835442.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HanjaPhoneticHangul
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HanjaPhoneticHangul");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HanjaPhoneticHangul", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld()
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", findText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase, matchWholeWord);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase, matchWholeWord, matchWildcards);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return Factory.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834930.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ClearFormatting()
		{
			 Factory.ExecuteMethod(this, "ClearFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194281.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void SetAllFuzzyOptions()
		{
			 Factory.ExecuteMethod(this, "SetAllFuzzyOptions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public void ClearAllFuzzyOptions()
		{
			 Factory.ExecuteMethod(this, "ClearAllFuzzyOptions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute()
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", findText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", findText, matchCase);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", findText, matchCase, matchWholeWord);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", findText, matchCase, matchWholeWord, matchWildcards);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		/// <param name="hanjaPhoneticHangul">optional object hanjaPhoneticHangul</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct, object hanjaPhoneticHangul)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace, ignorePunct, hanjaPhoneticHangul });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", findText);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor, textColor);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor, textColor, matchCase);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		/// <param name="textColor">optional object textColor</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="matchByte">optional object matchByte</param>
		/// <param name="matchFuzzy">optional object matchFuzzy</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct)
		{
			return Factory.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace, ignorePunct });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834830.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool ClearHitHighlight()
		{
			return Factory.ExecuteBoolMethodGet(this, "ClearHitHighlight");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		/// <param name="ignorePunct">optional object ignorePunct</param>
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace, object ignorePunct)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase, ignoreSpace, ignorePunct });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007()
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", findText);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase, matchWholeWord);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase, matchWholeWord, matchWildcards);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		/// <param name="matchWholeWord">optional object matchWholeWord</param>
		/// <param name="matchWildcards">optional object matchWildcards</param>
		/// <param name="matchSoundsLike">optional object matchSoundsLike</param>
		/// <param name="matchAllWordForms">optional object matchAllWordForms</param>
		/// <param name="forward">optional object forward</param>
		/// <param name="wrap">optional object wrap</param>
		/// <param name="format">optional object format</param>
		/// <param name="replaceWith">optional object replaceWith</param>
		/// <param name="replace">optional object replace</param>
		/// <param name="matchKashida">optional object matchKashida</param>
		/// <param name="matchDiacritics">optional object matchDiacritics</param>
		/// <param name="matchAlefHamza">optional object matchAlefHamza</param>
		/// <param name="matchControl">optional object matchControl</param>
		/// <param name="matchPrefix">optional object matchPrefix</param>
		/// <param name="matchSuffix">optional object matchSuffix</param>
		/// <param name="matchPhrase">optional object matchPhrase</param>
		/// <param name="ignoreSpace">optional object ignoreSpace</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace)
		{
			return Factory.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase, ignoreSpace });
		}

		#endregion

		#pragma warning restore
	}
}
