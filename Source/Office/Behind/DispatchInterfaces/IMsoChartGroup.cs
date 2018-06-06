using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IMsoChartGroup 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IMsoChartGroup : COMObject, NetOffice.OfficeApi.IMsoChartGroup
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
                    _type = typeof(IMsoChartGroup);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMsoChartGroup() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 AxisGroup
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "AxisGroup");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AxisGroup", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoDownBars DownBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDownBars>(this, "DownBars", typeof(NetOffice.OfficeApi.IMsoDownBars));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoDropLines DropLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDropLines>(this, "DropLines", typeof(NetOffice.OfficeApi.IMsoDropLines));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoHiLoLines HiLoLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoHiLoLines>(this, "HiLoLines", typeof(NetOffice.OfficeApi.IMsoHiLoLines));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Index
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Index");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object RadarAxisLabels
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "RadarAxisLabels");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoSeriesLines SeriesLines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoSeriesLines>(this, "SeriesLines", typeof(NetOffice.OfficeApi.IMsoSeriesLines));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoUpBars UpBars
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoUpBars>(this, "UpBars", typeof(NetOffice.OfficeApi.IMsoUpBars));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlSizeRepresents SizeRepresents
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlSizeRepresents>(this, "SizeRepresents");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SizeRepresents", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlChartSplitType SplitType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartSplitType>(this, "SplitType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SplitType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object SeriesCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "SeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object SeriesCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "SeriesCollection");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual object CategoryCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "CategoryCollection", index);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public virtual object CategoryCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "CategoryCollection");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual object FullCategoryCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection", index);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public virtual object FullCategoryCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "FullCategoryCollection");
        }

        #endregion

        #pragma warning restore
    }
}
