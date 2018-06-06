using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// IMsoChart
    /// </summary>
    [SyntaxBypass]
    public class IMsoChart_ : COMObject
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IMsoChart_() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ChartGroups(object pvarIndex, object varIgallery)
        {
            return Factory.ExecuteReferencePropertyGet(this, "ChartGroups", pvarIndex, varIgallery);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        public virtual object ChartGroups(object pvarIndex, object varIgallery)
        {
            return get_ChartGroups(pvarIndex, varIgallery);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ChartGroups(object pvarIndex)
        {
            return Factory.ExecuteReferencePropertyGet(this, "ChartGroups", pvarIndex);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        public virtual object ChartGroups(object pvarIndex)
        {
            return get_ChartGroups(pvarIndex);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object axisType, object axisGroup)
        {
            return Factory.ExecuteVariantPropertyGet(this, "HasAxis", axisType, axisGroup);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object axisType, object axisGroup, object value)
        {
            Factory.ExecutePropertySet(this, "HasAxis", axisType, axisGroup, value);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object axisType, object axisGroup)
        {
            return get_HasAxis(axisType, axisGroup);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object axisType)
        {
            return Factory.ExecuteVariantPropertyGet(this, "HasAxis", axisType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object axisType, object value)
        {
            Factory.ExecutePropertySet(this, "HasAxis", axisType, value);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object axisType)
        {
            return get_HasAxis(axisType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoWalls get_Walls(object fBackWall)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "Walls", typeof(NetOffice.OfficeApi.IMsoWalls), fBackWall);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Walls
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Walls")]
        public virtual NetOffice.OfficeApi.IMsoWalls Walls(object fBackWall)
        {
            return get_Walls(fBackWall);
        }

        #endregion

        #region Methods

        #endregion

    }

    /// <summary>
    /// DispatchInterface IMsoChart 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IMsoChart : NetOffice.OfficeApi.Behind.IMsoChart_, NetOffice.OfficeApi.IMsoChart
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
                    _type = typeof(IMsoChart);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IMsoChart() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
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
        [BaseResult]
        public virtual NetOffice.OfficeApi.IMsoChartTitle ChartTitle
        {
            get
            {
                return Factory.ExecuteBaseReferencePropertyGet<NetOffice.OfficeApi.IMsoChartTitle>(this, "ChartTitle");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 DepthPercent
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "DepthPercent");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "DepthPercent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Elevation
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Elevation");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Elevation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 GapDepth
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "GapDepth");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "GapDepth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 HeightPercent
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "HeightPercent");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HeightPercent", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Perspective
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Perspective");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Perspective", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object RightAngleAxes
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "RightAngleAxes");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "RightAngleAxes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Rotation
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Rotation");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Rotation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ProtectData
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectData");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectData", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ProtectFormatting
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectFormatting");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectFormatting", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ProtectGoalSeek
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectGoalSeek");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectGoalSeek", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ProtectSelection
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectSelection");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ProtectChartObjects
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectChartObjects");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectChartObjects", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object ChartGroups
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "ChartGroups");
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
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoCorners Corners
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoCorners>(this, "Corners", typeof(NetOffice.OfficeApi.IMsoCorners));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlChartType ChartType
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartType>(this, "ChartType");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ChartType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasDataTable
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasDataTable");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasDataTable", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlRowCol PlotBy
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlRowCol>(this, "PlotBy");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "PlotBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasLegend
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasLegend");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasLegend", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoLegend Legend
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoLegend>(this, "Legend", typeof(NetOffice.OfficeApi.IMsoLegend));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object HasAxis
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "HasAxis");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "HasAxis", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoWalls Walls
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "Walls", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoFloor Floor
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoFloor>(this, "Floor", typeof(NetOffice.OfficeApi.IMsoFloor));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoPlotArea PlotArea
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoPlotArea>(this, "PlotArea", typeof(NetOffice.OfficeApi.IMsoPlotArea));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool PlotVisibleOnly
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "PlotVisibleOnly");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "PlotVisibleOnly", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoChartArea ChartArea
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartArea>(this, "ChartArea", typeof(NetOffice.OfficeApi.IMsoChartArea));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool AutoScaling
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "AutoScaling");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "AutoScaling", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoDataTable DataTable
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDataTable>(this, "DataTable", typeof(NetOffice.OfficeApi.IMsoDataTable));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlBarShape BarShape
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlBarShape>(this, "BarShape");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BarShape", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoWalls SideWall
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "SideWall", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoWalls BackWall
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "BackWall", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ChartStyle
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ChartStyle");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ChartStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object PivotLayout
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "PivotLayout");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasPivotFields
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasPivotFields");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "HasPivotFields", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ShowDataLabelsOverMaximum
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowDataLabelsOverMaximum");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowDataLabelsOverMaximum", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public virtual object Selection
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Selection");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoChartData ChartData
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartData>(this, "ChartData", typeof(NetOffice.OfficeApi.IMsoChartData));
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Shapes Shapes
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.Shapes>(this, "Shapes", typeof(NetOffice.OfficeApi.Shapes));
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
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Area3DGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Area3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Bar3DGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Bar3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Column3DGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Column3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Line3DGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Line3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Pie3DGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Pie3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup SurfaceGroup
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "SurfaceGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual bool ShowReportFilterFieldButtons
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowReportFilterFieldButtons");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowReportFilterFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual bool ShowLegendFieldButtons
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowLegendFieldButtons");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowLegendFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual bool ShowAxisFieldButtons
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowAxisFieldButtons");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowAxisFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual bool ShowValueFieldButtons
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowValueFieldButtons");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowValueFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual bool ShowAllFieldButtons
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ShowAllFieldButtons");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ShowAllFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual bool ProtectChartSheetFormatting
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "ProtectChartSheetFormatting");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProtectChartSheetFormatting", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlCategoryLabelLevel>(this, "CategoryLabelLevel");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "CategoryLabelLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlSeriesNameLevel SeriesNameLevel
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlSeriesNameLevel>(this, "SeriesNameLevel");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "SeriesNameLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasHiddenContent
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "HasHiddenContent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual object ChartColor
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ChartColor");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ChartColor", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void UnProtect(object password)
        {
            Factory.ExecuteMethod(this, "UnProtect", password);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void UnProtect()
        {
            Factory.ExecuteMethod(this, "UnProtect");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
        {
            Factory.ExecuteMethod(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect()
        {
            Factory.ExecuteMethod(this, "Protect");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect(object password)
        {
            Factory.ExecuteMethod(this, "Protect", password);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect(object password, object drawingObjects)
        {
            Factory.ExecuteMethod(this, "Protect", password, drawingObjects);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect(object password, object drawingObjects, object contents)
        {
            Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Protect(object password, object drawingObjects, object contents, object scenarios)
        {
            Factory.ExecuteMethod(this, "Protect", password, drawingObjects, contents, scenarios);
        }

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
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            Factory.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void _ApplyDataLabels()
        {
            Factory.ExecuteMethod(this, "_ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type)
        {
            Factory.ExecuteMethod(this, "_ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey)
        {
            Factory.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            Factory.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels()
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
        {
            Factory.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
        {
            Factory.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
        {
            Factory.ExecuteMethod(this, "ApplyCustomType", chartType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
        {
            Factory.ExecuteMethod(this, "GetChartElement", new object[] { x, y, elementID, arg1, arg2 });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetSourceData(string source, object plotBy)
        {
            Factory.ExecuteMethod(this, "SetSourceData", source, plotBy);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetSourceData(string source)
        {
            Factory.ExecuteMethod(this, "SetSourceData", source);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Axes(object type, object axisGroup)
        {
            return Factory.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Axes()
        {
            return Factory.ExecuteVariantMethodGet(this, "Axes");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Axes(object type)
        {
            return Factory.ExecuteVariantMethodGet(this, "Axes", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AutoFormat(Int32 rGallery, object varFormat)
        {
            Factory.ExecuteMethod(this, "AutoFormat", rGallery, varFormat);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void AutoFormat(Int32 rGallery)
        {
            Factory.ExecuteMethod(this, "AutoFormat", rGallery);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetBackgroundPicture(string bstr)
        {
            Factory.ExecuteMethod(this, "SetBackgroundPicture", bstr);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        /// <param name="varValueTitle">optional object varValueTitle</param>
        /// <param name="varExtraTitle">optional object varExtraTitle</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle, object varExtraTitle)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle, varExtraTitle });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard()
        {
            Factory.ExecuteMethod(this, "ChartWizard");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource)
        {
            Factory.ExecuteMethod(this, "ChartWizard", varSource);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery)
        {
            Factory.ExecuteMethod(this, "ChartWizard", varSource, varGallery);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat)
        {
            Factory.ExecuteMethod(this, "ChartWizard", varSource, varGallery, varFormat);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy)
        {
            Factory.ExecuteMethod(this, "ChartWizard", varSource, varGallery, varFormat, varPlotBy);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        /// <param name="varTitle">optional object varTitle</param>
        /// <param name="varCategoryTitle">optional object varCategoryTitle</param>
        /// <param name="varValueTitle">optional object varValueTitle</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle)
        {
            Factory.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        /// <param name="size">optional Int32 Size = 2</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format, object size)
        {
            Factory.ExecuteMethod(this, "CopyPicture", appearance, format, size);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void CopyPicture()
        {
            Factory.ExecuteMethod(this, "CopyPicture");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void CopyPicture(object appearance)
        {
            Factory.ExecuteMethod(this, "CopyPicture", appearance);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format)
        {
            Factory.ExecuteMethod(this, "CopyPicture", appearance, format);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        /// <param name="objType">Int32 objType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Evaluate(object varName, Int32 localeID, out Int32 objType)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true);
            objType = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(varName, localeID, objType);
            object returnItem = Invoker.MethodReturn(this, "Evaluate", paramsArray, modifiers);
            if ((null != returnItem) && (returnItem is MarshalByRefObject))
            {
                ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem, true);
                objType = (Int32)paramsArray[2];
                return newObject;
            }
            else
            {
                objType = (Int32)paramsArray[2];
                return returnItem;
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object _Evaluate(object varName, Int32 localeID)
        {
            return Factory.ExecuteVariantMethodGet(this, "_Evaluate", varName, localeID);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varType">optional object varType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Paste(object varType)
        {
            Factory.ExecuteMethod(this, "Paste", varType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Paste()
        {
            Factory.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        /// <param name="varInteractive">optional object varInteractive</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Export(string bstr, object varFilterName, object varInteractive)
        {
            return Factory.ExecuteBoolMethodGet(this, "Export", bstr, varFilterName, varInteractive);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Export(string bstr)
        {
            return Factory.ExecuteBoolMethodGet(this, "Export", bstr);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Export(string bstr, object varFilterName)
        {
            return Factory.ExecuteBoolMethodGet(this, "Export", bstr, varFilterName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="varName">object varName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetDefaultChart(object varName)
        {
            Factory.ExecuteMethod(this, "SetDefaultChart", varName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyChartTemplate(string bstrFileName)
        {
            Factory.ExecuteMethod(this, "ApplyChartTemplate", bstrFileName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SaveChartTemplate(string bstrFileName)
        {
            Factory.ExecuteMethod(this, "SaveChartTemplate", bstrFileName);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ClearToMatchStyle()
        {
            Factory.ExecuteMethod(this, "ClearToMatchStyle");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void RefreshPivotTable()
        {
            Factory.ExecuteMethod(this, "RefreshPivotTable");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        /// <param name="varChartType">optional object varChartType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout, object varChartType)
        {
            Factory.ExecuteMethod(this, "ApplyLayout", layout, varChartType);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout)
        {
            Factory.ExecuteMethod(this, "ApplyLayout", layout);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void Refresh()
        {
            Factory.ExecuteMethod(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="rHS">NetOffice.OfficeApi.Enums.MsoChartElementType rHS</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType rHS)
        {
            Factory.ExecuteMethod(this, "SetElement", rHS);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object AreaGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "AreaGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object AreaGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "AreaGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object BarGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "BarGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object BarGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "BarGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object ColumnGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "ColumnGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object ColumnGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "ColumnGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object LineGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "LineGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object LineGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "LineGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object PieGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "PieGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object PieGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "PieGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object DoughnutGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object DoughnutGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "DoughnutGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object RadarGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "RadarGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object RadarGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "RadarGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object XYGroups(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "XYGroups", index);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object XYGroups()
        {
            return Factory.ExecuteVariantMethodGet(this, "XYGroups");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object Delete()
        {
            return Factory.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object Copy()
        {
            return Factory.ExecuteVariantMethodGet(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object Select(object replace)
        {
            return Factory.ExecuteVariantMethodGet(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual object Select()
        {
            return Factory.ExecuteVariantMethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 15, 16)]
        public virtual object FullSeriesCollection(object index)
        {
            return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 15, 16)]
        public virtual object FullSeriesCollection()
        {
            return Factory.ExecuteVariantMethodGet(this, "FullSeriesCollection");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 15, 16)]
        public virtual void DeleteHiddenContent()
        {
            Factory.ExecuteMethod(this, "DeleteHiddenContent");
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual void ClearToMatchColorStyle()
        {
            Factory.ExecuteMethod(this, "ClearToMatchColorStyle");
        }

        #endregion

        #pragma warning restore
    }
}
