using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface ListDataFormat 
	/// SupportByVersion Excel, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839751.aspx </remarks>
	[SupportByVersion("Excel", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ListDataFormat : COMObject, NetOffice.ExcelApi.ListDataFormat
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
                    _type = typeof(ListDataFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ListDataFormat() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834648.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841124.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834464.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlListDataType _Default
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlListDataType>(this, "_Default");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838807.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public object Choices
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Choices");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197836.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 DecimalPlaces
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "DecimalPlaces");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198254.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public object DefaultValue
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "DefaultValue");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196341.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool IsPercent
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsPercent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836203.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 lcid
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "lcid");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838252.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public Int32 MaxCharacters
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxCharacters");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821665.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public object MaxNumber
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MaxNumber");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836462.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public object MinNumber
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MinNumber");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839183.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool Required
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Required");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836844.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlListDataType Type
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlListDataType>(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836449.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool ReadOnly
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ReadOnly");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838058.aspx </remarks>
		[SupportByVersion("Excel", 11,12,14,15,16)]
		public bool AllowFillIn
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AllowFillIn");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


