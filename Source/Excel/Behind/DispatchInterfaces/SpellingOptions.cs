using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface SpellingOptions 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196915.aspx </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SpellingOptions : COMObject, NetOffice.ExcelApi.SpellingOptions
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
                    _type = typeof(SpellingOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SpellingOptions() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196429.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual Int32 DictLang
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DictLang");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DictLang", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835269.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual string UserDict
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "UserDict");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UserDict", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194482.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool IgnoreCaps
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreCaps");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreCaps", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840970.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool SuggestMainOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SuggestMainOnly");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SuggestMainOnly", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822188.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool IgnoreMixedDigits
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreMixedDigits");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreMixedDigits", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196363.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool IgnoreFileNames
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IgnoreFileNames");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IgnoreFileNames", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820788.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool GermanPostReform
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "GermanPostReform");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GermanPostReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836839.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool KoreanCombineAux
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KoreanCombineAux");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KoreanCombineAux", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837047.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool KoreanUseAutoChangeList
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KoreanUseAutoChangeList");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KoreanUseAutoChangeList", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838858.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool KoreanProcessCompound
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "KoreanProcessCompound");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "KoreanProcessCompound", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838254.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlHebrewModes HebrewModes
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlHebrewModes>(this, "HebrewModes");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "HebrewModes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193603.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlArabicModes ArabicModes
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlArabicModes>(this, "ArabicModes");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ArabicModes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193794.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual bool ArabicStrictAlefHamza
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ArabicStrictAlefHamza");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ArabicStrictAlefHamza", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835890.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual bool ArabicStrictFinalYaa
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ArabicStrictFinalYaa");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ArabicStrictFinalYaa", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841131.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual bool ArabicStrictTaaMarboota
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ArabicStrictTaaMarboota");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ArabicStrictTaaMarboota", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834631.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual bool RussianStrictE
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RussianStrictE");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RussianStrictE", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193307.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlSpanishModes SpanishModes
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSpanishModes>(this, "SpanishModes");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SpanishModes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822381.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlPortugueseReform PortugalReform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPortugueseReform>(this, "PortugalReform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "PortugalReform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839055.aspx </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlPortugueseReform BrazilReform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPortugueseReform>(this, "BrazilReform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BrazilReform", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}

