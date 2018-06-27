using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.PowerPointApi.Behind
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public class Chart_ : COMObject, NetOffice.PowerPointApi.Chart_
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object index1, object index2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", index1, index2);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object index1, object index2, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", index1, index2, value);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object index1, object index2)
        {
            return get_HasAxis(index1, index2);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object index1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", index1);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object index1, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", index1, value);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object index1)
        {
            return get_HasAxis(index1);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface Chart 
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744663.aspx </remarks>
    [SupportByVersion("PowerPoint", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Chart : Chart_, NetOffice.PowerPointApi.Chart
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
                    _contractType = typeof(NetOffice.PowerPointApi.Chart);
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746116.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744954.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745140.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744071.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlRowCol PlotBy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlRowCol>(this, "PlotBy");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PlotBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746809.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.DataTable DataTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.DataTable>(this, "DataTable", typeof(NetOffice.PowerPointApi.DataTable));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746790.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlBarShape BarShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlBarShape>(this, "BarShape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BarShape", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744381.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Walls SideWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "SideWall", typeof(NetOffice.PowerPointApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744079.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Walls BackWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "BackWall", typeof(NetOffice.PowerPointApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743954.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745647.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744089.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.ChartData ChartData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartData>(this, "ChartData", typeof(NetOffice.PowerPointApi.ChartData));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746059.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Shapes Shapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Shapes>(this, "Shapes", typeof(NetOffice.PowerPointApi.Shapes));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746336.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup Area3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Area3DGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup Bar3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Bar3DGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup Column3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Column3DGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup Line3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Line3DGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup Pie3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "Pie3DGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartGroup SurfaceGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartGroup>(this, "SurfaceGroup", typeof(NetOffice.PowerPointApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745066.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Application>(this, "Application", typeof(NetOffice.PowerPointApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744513.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744327.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.ChartArea ChartArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartArea>(this, "ChartArea", typeof(NetOffice.PowerPointApi.ChartArea));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743961.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.ChartTitle ChartTitle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartTitle>(this, "ChartTitle", typeof(NetOffice.PowerPointApi.ChartTitle));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.Corners Corners
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Corners>(this, "Corners", typeof(NetOffice.PowerPointApi.Corners));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746755.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745600.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745750.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745846.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Floor Floor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Floor>(this, "Floor", typeof(NetOffice.PowerPointApi.Floor));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746511.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746650.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743935.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746534.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745241.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744151.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Legend Legend
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Legend>(this, "Legend", typeof(NetOffice.PowerPointApi.Legend));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744105.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Name", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743957.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746093.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.PlotArea PlotArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.PlotArea>(this, "PlotArea", typeof(NetOffice.PowerPointApi.PlotArea));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745749.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744814.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745024.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 Subtype
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Subtype");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Subtype", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746542.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Walls Walls
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Walls>(this, "Walls", typeof(NetOffice.PowerPointApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745294.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.ChartFormat Format
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFormat>(this, "Format", typeof(NetOffice.PowerPointApi.ChartFormat));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745821.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743877.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744539.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746204.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744868.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746125.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string AlternativeText
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "AlternativeText");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "AlternativeText", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745833.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Title
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Title");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Title", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229264.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlCategoryLabelLevel>(this, "CategoryLabelLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CategoryLabelLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228519.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlSeriesNameLevel SeriesNameLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlSeriesNameLevel>(this, "SeriesNameLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SeriesNameLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasHiddenContent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasHiddenContent");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230443.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
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
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745773.aspx </remarks>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746151.aspx </remarks>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void GetChartElement(Int32 x, Int32 y, Int32 elementID, Int32 arg1, Int32 arg2)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "GetChartElement", new object[] { x, y, elementID, arg1, arg2 });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SetSourceData(string source, object plotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source, plotBy);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746759.aspx </remarks>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SetSourceData(string source)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void AutoFormat(Int32 gallery, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", gallery, format);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void AutoFormat(Int32 gallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", gallery);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745424.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SetBackgroundPicture(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBackgroundPicture", fileName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Paste(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste", type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746056.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745864.aspx </remarks>
        /// <param name="name">object name</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SetDefaultChart(object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultChart", name);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744899.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyChartTemplate(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyChartTemplate", fileName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744919.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SaveChartTemplate(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveChartTemplate", fileName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746785.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ClearToMatchStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchStyle");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        /// <param name="chartType">optional object chartType</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout, object chartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout, chartType);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745663.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745006.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Refresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object AreaGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object AreaGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object BarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object BarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object ColumnGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object ColumnGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object LineGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object LineGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object PieGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object PieGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object DoughnutGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object DoughnutGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object RadarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object RadarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object XYGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object XYGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void _ApplyDataLabels()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object legendKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.PowerPointApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void _ApplyDataLabels(object type, object legendKey, object autoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "_ApplyDataLabels", type, legendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.PowerPointApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Axes(object type, object axisGroup)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Axes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745216.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Axes(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object ChartGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ChartGroups", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744238.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object ChartGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ChartGroups");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        /// <param name="valueTitle">optional object valueTitle</param>
        /// <param name="extraTitle">optional object extraTitle</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery, format);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery, format, plotBy);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745899.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        /// <param name="categoryTitle">optional object categoryTitle</param>
        /// <param name="valueTitle">optional object valueTitle</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle });
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Copy(object before, object after)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", before, after);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745934.aspx </remarks>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Copy(object before)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", before);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
        /// <param name="size">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Size = 2</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format, object size)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format, size);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void CopyPicture()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void CopyPicture(object appearance)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745919.aspx </remarks>
        /// <param name="appearance">optional NetOffice.PowerPointApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.PowerPointApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745109.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Delete()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        /// <param name="interactive">optional object interactive</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual bool Export(string fileName, object filterName, object interactive)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName, filterName, interactive);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual bool Export(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744128.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual bool Export(string fileName, object filterName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName, filterName);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Select(object replace)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745013.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void Select()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Select");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object SeriesCollection(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745538.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object SeriesCollection()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746262.aspx </remarks>
        /// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetElement", element);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual object FullSeriesCollection(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FullSeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj228028.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual object FullSeriesCollection()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FullSeriesCollection");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual void DeleteHiddenContent()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteHiddenContent");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj227229.aspx </remarks>
        [SupportByVersion("PowerPoint", 15, 16)]
        public virtual void ClearToMatchColorStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchColorStyle");
        }

        #endregion

        #pragma warning restore
    }
}
