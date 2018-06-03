using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// DisplayFormat
    /// </summary>
    [SyntaxBypass]
    public class DisplayFormat_ : COMObject, NetOffice.ExcelApi.DisplayFormat_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public DisplayFormat_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.ExcelApi.Characters get_Characters(object start, object length)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters), start, length);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Excel", 14, 15, 16), Redirect("get_Characters")]
        public NetOffice.ExcelApi.Characters Characters(object start, object length)
        {
            return get_Characters(start, length);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx
        [SupportByVersion("Excel", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.ExcelApi.Characters get_Characters(object start)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters), start);
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Excel", 14, 15, 16), Redirect("get_Characters")]
        public NetOffice.ExcelApi.Characters Characters(object start)
        {
            return get_Characters(start);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface DisplayFormat 
    /// SupportByVersion Excel, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838863.aspx </remarks>
    [SupportByVersion("Excel", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class DisplayFormat : NetOffice.ExcelApi.Behind.DisplayFormat_, NetOffice.ExcelApi.DisplayFormat
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
                    _type = typeof(DisplayFormat);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public DisplayFormat() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838383.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822645.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194522.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195472.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Borders Borders
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Borders>(this, "Borders", typeof(NetOffice.ExcelApi.Borders));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197441.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Characters Characters
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Characters>(this, "Characters", typeof(NetOffice.ExcelApi.Characters));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836734.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Font Font
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Font>(this, "Font", typeof(NetOffice.ExcelApi.Font));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196382.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object Style
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Style");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821953.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object AddIndent
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "AddIndent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197010.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object FormulaHidden
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "FormulaHidden");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821609.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object HorizontalAlignment
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197906.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object IndentLevel
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "IndentLevel");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838619.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public NetOffice.ExcelApi.Interior Interior
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Interior>(this, "Interior", typeof(NetOffice.ExcelApi.Interior));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196267.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object Locked
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Locked");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837831.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object MergeCells
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "MergeCells");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198350.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object NumberFormat
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "NumberFormat");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193870.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object NumberFormatLocal
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "NumberFormatLocal");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840309.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object Orientation
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Orientation");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839759.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public Int32 ReadingOrder
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ReadingOrder");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834379.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object ShrinkToFit
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ShrinkToFit");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837378.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object VerticalAlignment
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "VerticalAlignment");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836493.aspx </remarks>
        [SupportByVersion("Excel", 14, 15, 16)]
        public object WrapText
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "WrapText");
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}

