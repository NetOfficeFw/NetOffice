using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface TextConnection 
	/// SupportByVersion Excel, 15, 16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228015.aspx </remarks>
	[SupportByVersion("Excel", 15, 16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class TextConnection : COMObject, NetOffice.ExcelApi.TextConnection
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
                    _type = typeof(TextConnection);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public TextConnection() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231033.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229623.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230690.aspx </remarks>
		[SupportByVersion("Excel", 15, 16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232073.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual object Connection
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Connection");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Connection", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231385.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileHeaderRow
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileHeaderRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileHeaderRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232041.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual object TextFileColumnDataTypes
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextFileColumnDataTypes");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextFileColumnDataTypes", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229203.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileCommaDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileCommaDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileCommaDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231987.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileConsecutiveDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileConsecutiveDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileConsecutiveDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231768.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual string TextFileDecimalSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileDecimalSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileDecimalSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228271.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual object TextFileFixedColumnWidths
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextFileFixedColumnWidths");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextFileFixedColumnWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227943.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual string TextFileOtherDelimiter
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileOtherDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileOtherDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227899.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlTextParsingType TextFileParseType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextParsingType>(this, "TextFileParseType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileParseType", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228433.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlPlatform TextFilePlatform
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlPlatform>(this, "TextFilePlatform");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFilePlatform", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227545.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFilePromptOnRefresh
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFilePromptOnRefresh");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFilePromptOnRefresh", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227524.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileSemicolonDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileSemicolonDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileSemicolonDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230951.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileSpaceDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileSpaceDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileSpaceDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231782.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual Int32 TextFileStartRow
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextFileStartRow");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileStartRow", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227355.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileTabDelimiter
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileTabDelimiter");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileTabDelimiter", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228275.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlTextQualifier TextFileTextQualifier
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextQualifier>(this, "TextFileTextQualifier");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileTextQualifier", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227507.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual string TextFileThousandsSeparator
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "TextFileThousandsSeparator");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileThousandsSeparator", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230624.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual bool TextFileTrailingMinusNumbers
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextFileTrailingMinusNumbers");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextFileTrailingMinusNumbers", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231092.aspx </remarks>
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.Enums.XlTextVisualLayoutType TextFileVisualLayout
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTextVisualLayoutType>(this, "TextFileVisualLayout");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextFileVisualLayout", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


