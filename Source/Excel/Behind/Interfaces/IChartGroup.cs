using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IChartGroup 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IChartGroup : COMObject, NetOffice.ExcelApi.IChartGroup
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
                    _type = typeof(IChartGroup);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IChartGroup() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
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
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAxisGroup AxisGroup
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisGroup>(this, "AxisGroup");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "AxisGroup", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 DoughnutHoleSize
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "DoughnutHoleSize");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DoughnutHoleSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.DownBars DownBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DownBars>(this, "DownBars", typeof(NetOffice.ExcelApi.DownBars));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.DropLines DropLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DropLines>(this, "DropLines", typeof(NetOffice.ExcelApi.DropLines));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 FirstSliceAngle
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "FirstSliceAngle");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FirstSliceAngle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 GapWidth
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "GapWidth");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "GapWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasDropLines
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasDropLines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasDropLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasHiLoLines
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasHiLoLines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasHiLoLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasRadarAxisLabels
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasRadarAxisLabels");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasRadarAxisLabels", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasSeriesLines
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasSeriesLines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasSeriesLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasUpDownBars
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasUpDownBars");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasUpDownBars", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.HiLoLines HiLoLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.HiLoLines>(this, "HiLoLines", typeof(NetOffice.ExcelApi.HiLoLines));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Index
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Overlap
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Overlap");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Overlap", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.TickLabels RadarAxisLabels
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TickLabels>(this, "RadarAxisLabels", typeof(NetOffice.ExcelApi.TickLabels));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.SeriesLines SeriesLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SeriesLines>(this, "SeriesLines", typeof(NetOffice.ExcelApi.SeriesLines));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 SubType
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SubType");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SubType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 Type
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Type");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.UpBars UpBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.UpBars>(this, "UpBars", typeof(NetOffice.ExcelApi.UpBars));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool VaryByCategories
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "VaryByCategories");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "VaryByCategories", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlSizeRepresents SizeRepresents
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlSizeRepresents>(this, "SizeRepresents");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SizeRepresents", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 BubbleScale
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "BubbleScale");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "BubbleScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ShowNegativeBubbles
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowNegativeBubbles");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowNegativeBubbles", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlChartSplitType SplitType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlChartSplitType>(this, "SplitType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SplitType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SplitValue
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "SplitValue");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "SplitValue", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 SecondPlotSize
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SecondPlotSize");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SecondPlotSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool Has3DShading
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Has3DShading");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Has3DShading", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SeriesCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object SeriesCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual object FullCategoryCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection", index);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object FullCategoryCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection");
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Excel", 15, 16)]
        public virtual object CategoryCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "CategoryCollection", index);
        }

        /// <summary>
        /// SupportByVersion Excel 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Excel", 15, 16)]
        public virtual object CategoryCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "CategoryCollection");
        }

        #endregion

        #pragma warning restore
    }
}

