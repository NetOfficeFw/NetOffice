using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface CellFormat 
	/// SupportByVersion Excel, 10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839805.aspx </remarks>
	[SupportByVersion("Excel", 10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class CellFormat : COMObject, NetOffice.ExcelApi.CellFormat
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
                    _type = typeof(CellFormat);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public CellFormat() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198158.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836509.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194095.aspx </remarks>
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
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835317.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Borders Borders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Borders>(this, "Borders", typeof(NetOffice.ExcelApi.Borders));
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Borders", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195695.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Font Font
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Font>(this, "Font", typeof(NetOffice.ExcelApi.Font));
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837616.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Interior Interior
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", typeof(NetOffice.ExcelApi.Interior));
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Interior", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820892.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object NumberFormat
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "NumberFormat");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "NumberFormat", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837569.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object NumberFormatLocal
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "NumberFormatLocal");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "NumberFormatLocal", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834695.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object AddIndent
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "AddIndent");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "AddIndent", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837837.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object IndentLevel
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "IndentLevel");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "IndentLevel", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196701.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object HorizontalAlignment
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839023.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object VerticalAlignment
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "VerticalAlignment");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841224.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object Orientation
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Orientation");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Orientation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836440.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object ShrinkToFit
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ShrinkToFit");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ShrinkToFit", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836142.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object WrapText
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "WrapText");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "WrapText", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822593.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object Locked
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Locked");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Locked", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821509.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object FormulaHidden
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "FormulaHidden");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "FormulaHidden", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194362.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual object MergeCells
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "MergeCells");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "MergeCells", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196018.aspx </remarks>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public virtual void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		#endregion

		#pragma warning restore
	}
}


