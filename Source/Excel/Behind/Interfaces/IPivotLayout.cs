using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// IPivotLayout
    /// </summary>
    [SyntaxBypass]
 	public class IPivotLayout_ : COMObject, NetOffice.ExcelApi.IPivotLayout_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPivotLayout_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_ColumnFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "ColumnFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_ColumnFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_ColumnFields")]
		public object ColumnFields(object index)
		{
			return get_ColumnFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_DataFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "DataFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_DataFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_DataFields")]
		public object DataFields(object index)
		{
			return get_DataFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_PageFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "PageFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_PageFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_PageFields")]
		public object PageFields(object index)
		{
			return get_PageFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_RowFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "RowFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_RowFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_RowFields")]
		public object RowFields(object index)
		{
			return get_RowFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_HiddenFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "HiddenFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_HiddenFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_HiddenFields")]
		public object HiddenFields(object index)
		{
			return get_HiddenFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_VisibleFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "VisibleFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_VisibleFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_VisibleFields")]
		public object VisibleFields(object index)
		{
			return get_VisibleFields(index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_PivotFields(object index)
		{
			return Factory.ExecuteReferencePropertyGet(this, "PivotFields", index);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Alias for get_PivotFields
		/// Unknown COM Proxy
		/// </summary>
		/// <param name="index">optional object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult, Redirect("get_PivotFields")]
		public object PivotFields(object index)
		{
			return get_PivotFields(index);
		}

		#endregion

		#region Methods

		#endregion
	}

    /// <summary>
    /// Interface IPivotLayout 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IPivotLayout : NetOffice.ExcelApi.Behind.IPivotLayout_, NetOffice.ExcelApi.IPivotLayout
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
                    _type = typeof(IPivotLayout);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IPivotLayout() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Application Application
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlCreator Creator
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object ColumnFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "ColumnFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object DataFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "DataFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object PageFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "PageFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object RowFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "RowFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object HiddenFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "HiddenFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object VisibleFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "VisibleFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public object PivotFields
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "PivotFields");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.CubeFields CubeFields
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.CubeFields>(this, "CubeFields", typeof(NetOffice.ExcelApi.CubeFields));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.PivotCache PivotCache
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotCache>(this, "PivotCache", typeof(NetOffice.ExcelApi.PivotCache));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.PivotTable PivotTable
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotTable>(this, "PivotTable", typeof(NetOffice.ExcelApi.PivotTable));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public string InnerDetail
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "InnerDetail");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "InnerDetail", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        /// <param name="appendField">optional object appendField</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddFields(object rowFields, object columnFields, object pageFields, object appendField)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddFields", rowFields, columnFields, pageFields, appendField);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddFields()
        {
            return Factory.ExecuteInt32MethodGet(this, "AddFields");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddFields(object rowFields)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddFields", rowFields);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddFields(object rowFields, object columnFields)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddFields", rowFields, columnFields);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="rowFields">optional object rowFields</param>
        /// <param name="columnFields">optional object columnFields</param>
        /// <param name="pageFields">optional object pageFields</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 AddFields(object rowFields, object columnFields, object pageFields)
        {
            return Factory.ExecuteInt32MethodGet(this, "AddFields", rowFields, columnFields, pageFields);
        }

        #endregion

        #pragma warning restore
    }
}

