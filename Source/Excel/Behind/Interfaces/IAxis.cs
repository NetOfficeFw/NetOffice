using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
    /// <summary>
    /// Interface IAxis 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IAxis : COMObject, NetOffice.ExcelApi.IAxis
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
                    _type = typeof(IAxis);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IAxis() : base()
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
        public virtual bool AxisBetweenCategories
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "AxisBetweenCategories");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AxisBetweenCategories", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAxisGroup AxisGroup
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisGroup>(this, "AxisGroup");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.AxisTitle AxisTitle
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.AxisTitle>(this, "AxisTitle", typeof(NetOffice.ExcelApi.AxisTitle));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Border Border
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Border>(this, "Border", typeof(NetOffice.ExcelApi.Border));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object CategoryNames
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "CategoryNames");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "CategoryNames", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAxisCrosses Crosses
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisCrosses>(this, "Crosses");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Crosses", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double CrossesAt
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "CrossesAt");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "CrossesAt", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasMajorGridlines
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasMajorGridlines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasMajorGridlines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasMinorGridlines
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasMinorGridlines");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasMinorGridlines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasTitle
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasTitle");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasTitle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Gridlines MajorGridlines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Gridlines>(this, "MajorGridlines", typeof(NetOffice.ExcelApi.Gridlines));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTickMark MajorTickMark
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickMark>(this, "MajorTickMark");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MajorTickMark", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double MajorUnit
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "MajorUnit");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MajorUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MajorUnitIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "MajorUnitIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MajorUnitIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double MaximumScale
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "MaximumScale");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MaximumScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MaximumScaleIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "MaximumScaleIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MaximumScaleIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double MinimumScale
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "MinimumScale");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MinimumScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MinimumScaleIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "MinimumScaleIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MinimumScaleIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Gridlines MinorGridlines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Gridlines>(this, "MinorGridlines", typeof(NetOffice.ExcelApi.Gridlines));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTickMark MinorTickMark
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickMark>(this, "MinorTickMark");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MinorTickMark", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double MinorUnit
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "MinorUnit");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MinorUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool MinorUnitIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "MinorUnitIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "MinorUnitIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool ReversePlotOrder
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ReversePlotOrder");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ReversePlotOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlScaleType ScaleType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlScaleType>(this, "ScaleType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ScaleType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTickLabelPosition TickLabelPosition
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTickLabelPosition>(this, "TickLabelPosition");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TickLabelPosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.TickLabels TickLabels
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.TickLabels>(this, "TickLabels", typeof(NetOffice.ExcelApi.TickLabels));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 TickLabelSpacing
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "TickLabelSpacing");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TickLabelSpacing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 TickMarkSpacing
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "TickMarkSpacing");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TickMarkSpacing", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlAxisType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlAxisType>(this, "Type");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTimeUnit BaseUnit
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "BaseUnit");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BaseUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool BaseUnitIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "BaseUnitIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "BaseUnitIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTimeUnit MajorUnitScale
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "MajorUnitScale");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MajorUnitScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlTimeUnit MinorUnitScale
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlTimeUnit>(this, "MinorUnitScale");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MinorUnitScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCategoryType CategoryType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCategoryType>(this, "CategoryType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "CategoryType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double Left
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Left");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double Top
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Top");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double Width
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double Height
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlDisplayUnit DisplayUnit
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlDisplayUnit>(this, "DisplayUnit");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "DisplayUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Double DisplayUnitCustom
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "DisplayUnitCustom");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DisplayUnitCustom", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual bool HasDisplayUnitLabel
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasDisplayUnitLabel");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasDisplayUnitLabel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.DisplayUnitLabel DisplayUnitLabel
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.DisplayUnitLabel>(this, "DisplayUnitLabel", typeof(NetOffice.ExcelApi.DisplayUnitLabel));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual Double LogBase
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "LogBase");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "LogBase", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual bool TickLabelSpacingIsAuto
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "TickLabelSpacingIsAuto");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TickLabelSpacingIsAuto", value);
            }
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.ChartFormat Format
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ChartFormat>(this, "Format", typeof(NetOffice.ExcelApi.ChartFormat));
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Delete()
        {
            return Factory.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual object Select()
        {
            return Factory.ExecuteVariantMethodGet(this, "Select");
        }

        #endregion

        #pragma warning restore
    }
}

