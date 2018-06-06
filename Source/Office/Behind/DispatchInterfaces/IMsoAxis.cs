using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IMsoAxis 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IMsoAxis : COMObject, NetOffice.OfficeApi.IMsoAxis
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
                    _type = typeof(IMsoAxis);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMsoAxis() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlAxisGroup>(this, "AxisGroup");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoAxisTitle AxisTitle
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoAxisTitle>(this, "AxisTitle", typeof(NetOffice.OfficeApi.IMsoAxisTitle));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlAxisCrosses Crosses
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlAxisCrosses>(this, "Crosses");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Crosses", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.GridLines MajorGridlines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GridLines>(this, "MajorGridlines", typeof(NetOffice.OfficeApi.GridLines));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTickMark MajorTickMark
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTickMark>(this, "MajorTickMark");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MajorTickMark", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.GridLines MinorGridlines
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.GridLines>(this, "MinorGridlines", typeof(NetOffice.OfficeApi.GridLines));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTickMark MinorTickMark
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTickMark>(this, "MinorTickMark");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MinorTickMark", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlScaleType ScaleType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlScaleType>(this, "ScaleType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ScaleType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTickLabelPosition TickLabelPosition
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTickLabelPosition>(this, "TickLabelPosition");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "TickLabelPosition", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoTickLabels TickLabels
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoTickLabels>(this, "TickLabels", typeof(NetOffice.OfficeApi.IMsoTickLabels));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlAxisType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlAxisType>(this, "Type");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTimeUnit BaseUnit
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTimeUnit>(this, "BaseUnit");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BaseUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTimeUnit MajorUnitScale
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTimeUnit>(this, "MajorUnitScale");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MajorUnitScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlTimeUnit MinorUnitScale
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlTimeUnit>(this, "MinorUnitScale");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "MinorUnitScale", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlCategoryType CategoryType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlCategoryType>(this, "CategoryType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "CategoryType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Double Left
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Left");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Double Top
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Top");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Double Width
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Double Height
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlDisplayUnit DisplayUnit
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlDisplayUnit>(this, "DisplayUnit");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "DisplayUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoDisplayUnitLabel DisplayUnitLabel
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDisplayUnitLabel>(this, "DisplayUnitLabel", typeof(NetOffice.OfficeApi.IMsoDisplayUnitLabel));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoBorder Border
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoBorder>(this, "Border", typeof(NetOffice.OfficeApi.IMsoBorder));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoChartFormat Format
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartFormat>(this, "Format", typeof(NetOffice.OfficeApi.IMsoChartFormat));
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
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Delete()
        {
            return Factory.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Select()
        {
            return Factory.ExecuteVariantMethodGet(this, "Select");
        }

        #endregion

        #pragma warning restore
    }
}
