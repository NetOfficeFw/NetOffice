using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface Interior 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196598.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Interior : COMObject, NetOffice.ExcelApi.Interior
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
                    _type = typeof(Interior);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public Interior() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839431.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840138.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194389.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840499.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object Color
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Color");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Color", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822603.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object ColorIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ColorIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197910.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object InvertIfNegative
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "InvertIfNegative");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "InvertIfNegative", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835883.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object Pattern
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Pattern");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Pattern", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197513.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object PatternColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PatternColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PatternColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840377.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public virtual object PatternColorIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PatternColorIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PatternColorIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff820778.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual object ThemeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ThemeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ThemeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197557.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual object TintAndShade
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TintAndShade");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TintAndShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839013.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual object PatternThemeColor
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PatternThemeColor");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PatternThemeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192986.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual object PatternTintAndShade
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "PatternTintAndShade");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "PatternTintAndShade", value);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195190.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public virtual object Gradient
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Gradient");
			}
		}

		#endregion

		#region Methods

		#endregion

		#pragma warning restore
	}
}


