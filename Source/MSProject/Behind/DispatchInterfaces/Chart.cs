using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi.Behind
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public class Chart_ : COMObject, NetOffice.MSProjectApi.Chart_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Chart_() : base()
        {
        }

        #endregion

        #region Properties
        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("MSProject", 11), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ChartGroups(object pvarIndex, object varIgallery)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartGroups", pvarIndex, varIgallery);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        /// <param name="varIgallery">optional object varIgallery</param>
        [SupportByVersion("MSProject", 11), ProxyResult, Redirect("get_ChartGroups")]
        public virtual object ChartGroups(object pvarIndex, object varIgallery)
        {
            return get_ChartGroups(pvarIndex, varIgallery);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("MSProject", 11), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ChartGroups(object pvarIndex)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartGroups", pvarIndex);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="pvarIndex">optional object pvarIndex</param>
        [SupportByVersion("MSProject", 11), ProxyResult, Redirect("get_ChartGroups")]
        public virtual object ChartGroups(object pvarIndex)
        {
            return get_ChartGroups(pvarIndex);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object axisType, object axisGroup)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", axisType, axisGroup);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object axisType, object axisGroup, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", axisType, axisGroup, value);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="axisGroup">optional object axisGroup</param>
        [SupportByVersion("MSProject", 11), Redirect("get_HasAxis")]
        public virtual object HasAxis(object axisType, object axisGroup)
        {
            return get_HasAxis(axisType, axisGroup);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object axisType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", axisType);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object axisType, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", axisType, value);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_HasAxis
        /// </summary>
        /// <param name="axisType">optional object axisType</param>
        [SupportByVersion("MSProject", 11), Redirect("get_HasAxis")]
        public virtual object HasAxis(object axisType)
        {
            return get_HasAxis(axisType);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoWalls get_Walls(object fBackWall)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "Walls", typeof(NetOffice.OfficeApi.IMsoWalls), fBackWall);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Alias for get_Walls
        /// </summary>
        /// <param name="fBackWall">optional bool fBackWall</param>
        [SupportByVersion("MSProject", 11), Redirect("get_Walls")]
        public virtual NetOffice.OfficeApi.IMsoWalls Walls(object fBackWall)
        {
            return get_Walls(fBackWall);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Chart 
    /// SupportByVersion MSProject, 11
    /// </summary>
    [SupportByVersion("MSProject", 11)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Chart : Chart_, NetOffice.MSProjectApi.Chart
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.MSProjectApi.Chart);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(Chart);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public Chart() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool HasTitle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasTitle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasTitle", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [BaseResult]
        public virtual NetOffice.OfficeApi.IMsoChartTitle ChartTitle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBaseReferencePropertyGet<NetOffice.OfficeApi.IMsoChartTitle>(this, "ChartTitle");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 DepthPercent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "DepthPercent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DepthPercent", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 Elevation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Elevation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Elevation", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 GapDepth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GapDepth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GapDepth", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 HeightPercent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "HeightPercent");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HeightPercent", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 Perspective
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Perspective");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Perspective", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object RightAngleAxes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "RightAngleAxes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "RightAngleAxes", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object Rotation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Rotation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Rotation", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ProtectData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProtectData", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ProtectFormatting
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectFormatting");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProtectFormatting", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ProtectGoalSeek
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectGoalSeek");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProtectGoalSeek", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ProtectSelection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectSelection");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProtectSelection", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ProtectChartObjects
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ProtectChartObjects");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ProtectChartObjects", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        public virtual object ChartGroups
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartGroups");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 SubType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "SubType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "SubType", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 Type
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Type");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoCorners Corners
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoCorners>(this, "Corners", typeof(NetOffice.OfficeApi.IMsoCorners));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.Enums.XlChartType ChartType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartType>(this, "ChartType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "ChartType", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool HasDataTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasDataTable");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasDataTable", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.Enums.XlRowCol PlotBy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlRowCol>(this, "PlotBy");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PlotBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool HasLegend
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasLegend");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasLegend", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoLegend Legend
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoLegend>(this, "Legend", typeof(NetOffice.OfficeApi.IMsoLegend));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object HasAxis
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HasAxis", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoWalls Walls
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "Walls", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoFloor Floor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoFloor>(this, "Floor", typeof(NetOffice.OfficeApi.IMsoFloor));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoPlotArea PlotArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoPlotArea>(this, "PlotArea", typeof(NetOffice.OfficeApi.IMsoPlotArea));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool PlotVisibleOnly
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "PlotVisibleOnly");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PlotVisibleOnly", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoChartArea ChartArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartArea>(this, "ChartArea", typeof(NetOffice.OfficeApi.IMsoChartArea));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool AutoScaling
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "AutoScaling");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AutoScaling", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoDataTable DataTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoDataTable>(this, "DataTable", typeof(NetOffice.OfficeApi.IMsoDataTable));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.Enums.XlBarShape BarShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlBarShape>(this, "BarShape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BarShape", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoWalls SideWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "SideWall", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoWalls BackWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoWalls>(this, "BackWall", typeof(NetOffice.OfficeApi.IMsoWalls));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object ChartStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ChartStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ChartStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        public virtual object PivotLayout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PivotLayout");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasPivotFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasPivotFields");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasPivotFields", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowDataLabelsOverMaximum
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowDataLabelsOverMaximum");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowDataLabelsOverMaximum", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        public virtual object Selection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Selection");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoChartData ChartData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartData>(this, "ChartData", typeof(NetOffice.OfficeApi.IMsoChartData));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.OfficeApi.IMsoChartFormat Format
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartFormat>(this, "Format", typeof(NetOffice.OfficeApi.IMsoChartFormat));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual NetOffice.MSProjectApi.Shapes Shapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.MSProjectApi.Shapes>(this, "Shapes", typeof(NetOffice.MSProjectApi.Shapes));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("MSProject", 11), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Area3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Area3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Bar3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Bar3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Column3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Column3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Line3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Line3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup Pie3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "Pie3DGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OfficeApi.IMsoChartGroup SurfaceGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartGroup>(this, "SurfaceGroup", typeof(NetOffice.OfficeApi.IMsoChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowReportFilterFieldButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowReportFilterFieldButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowReportFilterFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowLegendFieldButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowLegendFieldButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowLegendFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowAxisFieldButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowAxisFieldButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowAxisFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowValueFieldButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowValueFieldButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowValueFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual bool ShowAllFieldButtons
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ShowAllFieldButtons");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ShowAllFieldButtons", value);
            }
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// Get/Set
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object ChartColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "ChartColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "ChartColor", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void UnProtect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UnProtect", password);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UnProtect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UnProtect");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        /// <param name="userInterfaceOnly">optional object userInterfaceOnly</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect(object password, object drawingObjects, object contents, object scenarios, object userInterfaceOnly)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", new object[] { password, drawingObjects, contents, scenarios, userInterfaceOnly });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect(object password)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect(object password, object drawingObjects)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, drawingObjects);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect(object password, object drawingObjects, object contents)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, drawingObjects, contents);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="password">optional object password</param>
        /// <param name="drawingObjects">optional object drawingObjects</param>
        /// <param name="contents">optional object contents</param>
        /// <param name="scenarios">optional object scenarios</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Protect(object password, object drawingObjects, object contents, object scenarios)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Protect", password, drawingObjects, contents, scenarios);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("MSProject", 11)]
        public virtual object SeriesCollection(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object SeriesCollection()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void _ApplyDataLabels()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void _ApplyDataLabels(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void _ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "GetChartElement", new object[] { x, y, elementID, arg1, arg2 });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void SetSourceData(string source, object plotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source, plotBy);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void SetSourceData(string source)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("MSProject", 11)]
        public virtual object Axes(object type, object axisGroup)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object Axes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object Axes(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void AutoFormat(Int32 rGallery, object varFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", rGallery, varFormat);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rGallery">Int32 rGallery</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void AutoFormat(Int32 rGallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", rGallery);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void SetBackgroundPicture(string bstr)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBackgroundPicture", bstr);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle, object varExtraTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle, varExtraTitle });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", varSource);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", varSource, varGallery);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", varSource, varGallery, varFormat);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", varSource, varGallery, varFormat, varPlotBy);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varSource">optional object varSource</param>
        /// <param name="varGallery">optional object varGallery</param>
        /// <param name="varFormat">optional object varFormat</param>
        /// <param name="varPlotBy">optional object varPlotBy</param>
        /// <param name="varCategoryLabels">optional object varCategoryLabels</param>
        /// <param name="varSeriesLabels">optional object varSeriesLabels</param>
        /// <param name="varHasLegend">optional object varHasLegend</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
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
        [SupportByVersion("MSProject", 11)]
        public virtual void ChartWizard(object varSource, object varGallery, object varFormat, object varPlotBy, object varCategoryLabels, object varSeriesLabels, object varHasLegend, object varTitle, object varCategoryTitle, object varValueTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { varSource, varGallery, varFormat, varPlotBy, varCategoryLabels, varSeriesLabels, varHasLegend, varTitle, varCategoryTitle, varValueTitle });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        /// <param name="size">optional Int32 Size = 2</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void CopyPicture(object appearance, object format, object size)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format, size);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void CopyPicture()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void CopyPicture(object appearance)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="appearance">optional Int32 Appearance = 1</param>
        /// <param name="format">optional Int32 Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void CopyPicture(object appearance, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        /// <param name="objType">Int32 objType</param>
        [SupportByVersion("MSProject", 11)]
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
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varName">object varName</param>
        /// <param name="localeID">Int32 localeID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object _Evaluate(object varName, Int32 localeID)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_Evaluate", varName, localeID);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varType">optional object varType</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void Paste(object varType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste", varType);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        /// <param name="varInteractive">optional object varInteractive</param>
        [SupportByVersion("MSProject", 11)]
        public virtual bool Export(string bstr, object varFilterName, object varInteractive)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", bstr, varFilterName, varInteractive);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual bool Export(string bstr)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", bstr);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstr">string bstr</param>
        /// <param name="varFilterName">optional object varFilterName</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual bool Export(string bstr, object varFilterName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", bstr, varFilterName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="varName">object varName</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void SetDefaultChart(object varName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultChart", varName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyChartTemplate(string bstrFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyChartTemplate", bstrFileName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="bstrFileName">string bstrFileName</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void SaveChartTemplate(string bstrFileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveChartTemplate", bstrFileName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual void ClearToMatchStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchStyle");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual void RefreshPivotTable()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshPivotTable");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        /// <param name="varChartType">optional object varChartType</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyLayout(Int32 layout, object varChartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout, varChartType);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void ApplyLayout(Int32 layout)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual void Refresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="rHS">NetOffice.OfficeApi.Enums.MsoChartElementType rHS</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType rHS)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetElement", rHS);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object AreaGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object AreaGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object BarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object BarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object ColumnGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object ColumnGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object LineGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object LineGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object PieGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object PieGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object DoughnutGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object DoughnutGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object RadarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object RadarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("MSProject", 11)]
        public virtual object XYGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups", index);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object XYGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object Delete()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual object Copy()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("MSProject", 11)]
        public virtual object Select(object replace)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual object Select()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        /// <param name="startDate">optional object startDate</param>
        /// <param name="finishDate">optional object finishDate</param>
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount, object startDate, object finishDate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField, safeArrayOfPjTimescaledData, timeScaleUnit, timescaleUnitCount, startDate, finishDate });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", task, timephased);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", task, timephased, groupName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", task, timephased, groupName, filterName);
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField, safeArrayOfPjTimescaledData });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField, safeArrayOfPjTimescaledData, timeScaleUnit });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField, safeArrayOfPjTimescaledData, timeScaleUnit, timescaleUnitCount });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        /// <param name="task">bool task</param>
        /// <param name="timephased">bool timephased</param>
        /// <param name="groupName">optional string GroupName = </param>
        /// <param name="filterName">optional string FilterName = </param>
        /// <param name="labelField">optional NetOffice.MSProjectApi.Enums.PjField LabelField = -1</param>
        /// <param name="outlineLevel">optional Int32 OutlineLevel = -1</param>
        /// <param name="safeArrayOfPjField">optional object safeArrayOfPjField</param>
        /// <param name="safeArrayOfPjTimescaledData">optional object safeArrayOfPjTimescaledData</param>
        /// <param name="timeScaleUnit">optional NetOffice.MSProjectApi.Enums.PjTimescaleUnit TimeScaleUnit = 4</param>
        /// <param name="timescaleUnitCount">optional Int32 TimescaleUnitCount = 1</param>
        /// <param name="startDate">optional object startDate</param>
        [CustomMethod]
        [SupportByVersion("MSProject", 11)]
        public virtual void UpdateChartData(bool task, bool timephased, object groupName, object filterName, object labelField, object outlineLevel, object safeArrayOfPjField, object safeArrayOfPjTimescaledData, object timeScaleUnit, object timescaleUnitCount, object startDate)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateChartData", new object[] { task, timephased, groupName, filterName, labelField, outlineLevel, safeArrayOfPjField, safeArrayOfPjTimescaledData, timeScaleUnit, timescaleUnitCount, startDate });
        }

        /// <summary>
        /// SupportByVersion MSProject 11
        /// </summary>
        [SupportByVersion("MSProject", 11)]
        public virtual void ClearToMatchColorStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchColorStyle");
        }

        #endregion

        #pragma warning restore
    }
}

