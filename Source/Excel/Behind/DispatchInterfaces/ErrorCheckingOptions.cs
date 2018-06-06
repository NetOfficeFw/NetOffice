using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface ErrorCheckingOptions 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841029.aspx </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ErrorCheckingOptions : COMObject, NetOffice.ExcelApi.ErrorCheckingOptions
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
                    _type = typeof(ErrorCheckingOptions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ErrorCheckingOptions() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196328.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839201.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821627.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197304.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool BackgroundChecking
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BackgroundChecking");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackgroundChecking", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197740.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlColorIndex IndicatorColorIndex
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlColorIndex>(this, "IndicatorColorIndex");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IndicatorColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841054.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool EvaluateToError
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EvaluateToError");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EvaluateToError", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840516.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool TextDate
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "TextDate");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "TextDate", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198182.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool NumberAsText
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "NumberAsText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "NumberAsText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835006.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool InconsistentFormula
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InconsistentFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InconsistentFormula", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837361.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool OmittedCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "OmittedCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OmittedCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193607.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool UnlockedFormulaCells
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UnlockedFormulaCells");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UnlockedFormulaCells", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196993.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual bool EmptyCellReferences
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "EmptyCellReferences");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "EmptyCellReferences", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836531.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public virtual bool ListDataValidation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ListDataValidation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListDataValidation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839047.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual bool InconsistentTableFormula
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "InconsistentTableFormula");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "InconsistentTableFormula", value);
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


