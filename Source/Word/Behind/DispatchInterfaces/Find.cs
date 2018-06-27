using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.WordApi;

namespace NetOffice.WordApi.Behind
{
	/// <summary>
	/// DispatchInterface Find 
	/// SupportByVersion Word, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839118.aspx </remarks>
	[SupportByVersion("Word", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Find : COMObject, NetOffice.WordApi.Find
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.WordApi.Find);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Find() : base()
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
		public virtual NetOffice.WordApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Application>(this, "Application", typeof(NetOffice.WordApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839624.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834556.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839325.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Forward
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Forward");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Forward", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822678.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Font Font
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Font>(this, "Font", typeof(NetOffice.WordApi.Font));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838143.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Found
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Found");
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845697.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchAllWordForms
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchAllWordForms");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchAllWordForms", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837923.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchCase
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchCase");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchCase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838695.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchWildcards
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchWildcards");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchWildcards", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821942.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchSoundsLike
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchSoundsLike");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchSoundsLike", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835745.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchWholeWord
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchWholeWord");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchWholeWord", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821682.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchFuzzy
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchFuzzy");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchFuzzy", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838094.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchByte
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchByte");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchByte", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836406.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.ParagraphFormat ParagraphFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ParagraphFormat>(this, "ParagraphFormat", typeof(NetOffice.WordApi.ParagraphFormat));
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteReferencePropertySet(this, "ParagraphFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual object Style
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Style");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838976.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual string Text
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837887.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageID
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageID");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageID", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821028.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 Highlight
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Highlight");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Highlight", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836618.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Replacement Replacement
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Replacement>(this, "Replacement", typeof(NetOffice.WordApi.Replacement));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197498.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Frame Frame
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Frame>(this, "Frame", typeof(NetOffice.WordApi.Frame));
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192810.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdFindWrap Wrap
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdFindWrap>(this, "Wrap");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Wrap", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834863.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Format
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Format");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Format", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195137.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageIDFarEast
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDFarEast");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageIDFarEast", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836860.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual NetOffice.WordApi.Enums.WdLanguageID LanguageIDOther
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.WdLanguageID>(this, "LanguageIDOther");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "LanguageIDOther", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821910.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool CorrectHangulEndings
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "CorrectHangulEndings");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "CorrectHangulEndings", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195417.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual Int32 NoProofing
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "NoProofing");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "NoProofing", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845200.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchKashida
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchKashida");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchKashida", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839133.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchDiacritics
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchDiacritics");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchDiacritics", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845597.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchAlefHamza
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchAlefHamza");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchAlefHamza", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194643.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool MatchControl
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchControl");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchControl", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191768.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool MatchPhrase
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchPhrase");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchPhrase", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197820.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool MatchPrefix
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchPrefix");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchPrefix", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839710.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool MatchSuffix
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "MatchSuffix");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MatchSuffix", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821316.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool IgnoreSpace
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IgnoreSpace");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IgnoreSpace", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194518.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool IgnorePunct
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IgnorePunct");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IgnorePunct", value);
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835442.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool HanjaPhoneticHangul
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HanjaPhoneticHangul");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HanjaPhoneticHangul", value);
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ExecuteOld()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ExecuteOld(object findText)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", findText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool ExecuteOld(object findText, object matchCase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase);
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase, matchWholeWord);
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", findText, matchCase, matchWholeWord, matchWildcards);
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
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
		public virtual bool ExecuteOld(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ExecuteOld", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834930.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ClearFormatting()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearFormatting");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194281.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void SetAllFuzzyOptions()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "SetAllFuzzyOptions");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838471.aspx </remarks>
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual void ClearAllFuzzyOptions()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "ClearAllFuzzyOptions");
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Execute()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute");
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Execute(object findText)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", findText);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193977.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 9,10,11,12,14,15,16)]
		public virtual bool Execute(object findText, object matchCase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", findText, matchCase);
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", findText, matchCase, matchWholeWord);
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", findText, matchCase, matchWholeWord, matchWildcards);
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics });
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
		public virtual bool Execute(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct, object hanjaPhoneticHangul)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace, ignorePunct, hanjaPhoneticHangul });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool HitHighlight(object findText)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", findText);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845691.aspx </remarks>
		/// <param name="findText">object findText</param>
		/// <param name="highlightColor">optional object highlightColor</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool HitHighlight(object findText, object highlightColor)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor);
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor, textColor);
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", findText, highlightColor, textColor, matchCase);
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace });
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
		public virtual bool HitHighlight(object findText, object highlightColor, object textColor, object matchCase, object matchWholeWord, object matchPrefix, object matchSuffix, object matchPhrase, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object matchByte, object matchFuzzy, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object ignoreSpace, object ignorePunct)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "HitHighlight", new object[]{ findText, highlightColor, textColor, matchCase, matchWholeWord, matchPrefix, matchSuffix, matchPhrase, matchWildcards, matchSoundsLike, matchAllWordForms, matchByte, matchFuzzy, matchKashida, matchDiacritics, matchAlefHamza, matchControl, ignoreSpace, ignorePunct });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834830.aspx </remarks>
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool ClearHitHighlight()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "ClearHitHighlight");
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace, object ignorePunct)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase, ignoreSpace, ignorePunct });
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool Execute2007()
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007");
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool Execute2007(object findText)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", findText);
		}

		/// <summary>
		/// SupportByVersion Word 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194658.aspx </remarks>
		/// <param name="findText">optional object findText</param>
		/// <param name="matchCase">optional object matchCase</param>
		[CustomMethod]
		[SupportByVersion("Word", 12,14,15,16)]
		public virtual bool Execute2007(object findText, object matchCase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase);
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase, matchWholeWord);
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", findText, matchCase, matchWholeWord, matchWildcards);
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase });
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
		public virtual bool Execute2007(object findText, object matchCase, object matchWholeWord, object matchWildcards, object matchSoundsLike, object matchAllWordForms, object forward, object wrap, object format, object replaceWith, object replace, object matchKashida, object matchDiacritics, object matchAlefHamza, object matchControl, object matchPrefix, object matchSuffix, object matchPhrase, object ignoreSpace)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Execute2007", new object[]{ findText, matchCase, matchWholeWord, matchWildcards, matchSoundsLike, matchAllWordForms, forward, wrap, format, replaceWith, replace, matchKashida, matchDiacritics, matchAlefHamza, matchControl, matchPrefix, matchSuffix, matchPhrase, ignoreSpace });
		}

		#endregion

		#pragma warning restore
	}
}


