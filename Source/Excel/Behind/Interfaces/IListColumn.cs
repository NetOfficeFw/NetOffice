using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IListColumn 
    /// SupportByVersion Excel, 11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IListColumn : COMObject, NetOffice.ExcelApi.IListColumn
    {
        #pragma warning disable

        #region Type Information

        /// <summary>        /// Instance Type
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
                    _type = typeof(IListColumn);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IListColumn() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16), ProxyResult]
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
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public string _Default
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "_Default");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.ListDataFormat ListDataFormat
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ListDataFormat>(this, "ListDataFormat", typeof(NetOffice.ExcelApi.ListDataFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public Int32 Index
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Range Range
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Range", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Enums.XlTotalsCalculation TotalsCalculation
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTotalsCalculation>(this, "TotalsCalculation");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TotalsCalculation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public NetOffice.ExcelApi.XPath XPath
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.XPath>(this, "XPath", typeof(NetOffice.ExcelApi.XPath));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public string SharePointFormula
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "SharePointFormula");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Range DataBodyRange
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "DataBodyRange", typeof(NetOffice.ExcelApi.Range));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public NetOffice.ExcelApi.Range Total
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Range>(this, "Total", typeof(NetOffice.ExcelApi.Range));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 11, 12, 14, 15, 16)]
        public Int32 Delete()
        {
            return Factory.ExecuteInt32MethodGet(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}

