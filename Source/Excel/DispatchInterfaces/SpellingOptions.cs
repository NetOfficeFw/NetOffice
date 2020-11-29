﻿using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// DispatchInterface SpellingOptions 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions"/> </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SpellingOptions : COMObject
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

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SpellingOptions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SpellingOptions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SpellingOptions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.DictLang"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public Int32 DictLang
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.UserDict"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public string UserDict
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.IgnoreCaps"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool IgnoreCaps
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.SuggestMainOnly"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool SuggestMainOnly
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.IgnoreMixedDigits"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool IgnoreMixedDigits
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.IgnoreFileNames"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool IgnoreFileNames
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.GermanPostReform"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool GermanPostReform
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.KoreanCombineAux"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool KoreanCombineAux
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.KoreanUseAutoChangeList"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool KoreanUseAutoChangeList
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.KoreanProcessCompound"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public bool KoreanProcessCompound
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.HebrewModes"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlHebrewModes HebrewModes
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.ArabicModes"/> </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlArabicModes ArabicModes
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.ArabicStrictAlefHamza"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ArabicStrictAlefHamza
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.ArabicStrictFinalYaa"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ArabicStrictFinalYaa
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.ArabicStrictTaaMarboota"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool ArabicStrictTaaMarboota
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.RussianStrictE"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public bool RussianStrictE
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.SpanishModes"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlSpanishModes SpanishModes
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.PortugalReform"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPortugueseReform PortugalReform
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.SpellingOptions.BrazilReform"/> </remarks>
		[SupportByVersion("Excel", 14,15,16)]
		public NetOffice.ExcelApi.Enums.XlPortugueseReform BrazilReform
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
