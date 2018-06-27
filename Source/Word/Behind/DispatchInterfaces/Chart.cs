using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.WordApi.Behind
{
    /// <summary>
    /// Chart
    /// </summary>
    [SyntaxBypass]
    public class Chart_ : COMObject
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public Chart_() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_ChartGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx
        /// Alias for get_ChartGroups
        /// Unknown COM Proxy
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult, Redirect("get_ChartGroups")]
        public virtual object ChartGroups(object index)
        {
            return get_ChartGroups(index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object index1, object index2)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", index1, index2);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object index1, object index2, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", index1, index2, value);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        /// <param name="index2">optional object index2</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object index1, object index2)
        {
            return get_HasAxis(index1, index2);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object get_HasAxis(object index1)
        {
            return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HasAxis", index1);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <param name="index1">optional object index1</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual void set_HasAxis(object index1, object value)
        {
            InvokerService.InvokeInternal.ExecutePropertySet(this, "HasAxis", index1, value);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Alias for get_HasAxis
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        /// <param name="index1">optional object index1</param>
        [SupportByVersion("Word", 14, 15, 16), Redirect("get_HasAxis")]
        public virtual object HasAxis(object index1)
        {
            return get_HasAxis(index1);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface Chart 
    /// SupportByVersion Word, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193446.aspx </remarks>
    [SupportByVersion("Word", 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class Chart : Chart_, NetOffice.WordApi.Chart
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
                    _contractType = typeof(NetOffice.WordApi.Chart);
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
        /// Stub Ctor, not indented to use
        /// </summary>
        public Chart() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191738.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196350.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191751.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.ChartTitle ChartTitle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartTitle>(this, "ChartTitle", typeof(NetOffice.WordApi.ChartTitle));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840907.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192611.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845244.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836594.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838954.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838938.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835465.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196216.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.XlDisplayBlanksAs DisplayBlanksAs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlDisplayBlanksAs>(this, "DisplayBlanksAs");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DisplayBlanksAs", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836391.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        public virtual object ChartGroups
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "ChartGroups");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.Corners Corners
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Corners>(this, "Corners", typeof(NetOffice.WordApi.Corners));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836334.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197158.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836380.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.XlRowCol PlotBy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlRowCol>(this, "PlotBy");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PlotBy", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845054.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836685.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Legend Legend
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Legend>(this, "Legend", typeof(NetOffice.WordApi.Legend));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836998.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840511.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Walls Walls
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "Walls", typeof(NetOffice.WordApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845855.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Floor Floor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Floor>(this, "Floor", typeof(NetOffice.WordApi.Floor));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194655.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.PlotArea PlotArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.PlotArea>(this, "PlotArea", typeof(NetOffice.WordApi.PlotArea));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196388.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836658.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.ChartArea ChartArea
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartArea>(this, "ChartArea", typeof(NetOffice.WordApi.ChartArea));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823268.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191967.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.DataTable DataTable
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.DataTable>(this, "DataTable", typeof(NetOffice.WordApi.DataTable));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839500.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Enums.XlBarShape BarShape
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlBarShape>(this, "BarShape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BarShape", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839285.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Walls SideWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "SideWall", typeof(NetOffice.WordApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193753.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.Walls BackWall
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.Walls>(this, "BackWall", typeof(NetOffice.WordApi.Walls));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195916.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836370.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        public virtual object PivotLayout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PivotLayout");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193871.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838941.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual NetOffice.WordApi.ChartData ChartData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartData>(this, "ChartData", typeof(NetOffice.WordApi.ChartData));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837462.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        public virtual object Shapes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Shapes");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835828.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16), ProxyResult]
        public virtual object Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192382.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual Int32 Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup Area3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Area3DGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup Bar3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Bar3DGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup Column3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Column3DGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup Line3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Line3DGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup Pie3DGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "Pie3DGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.WordApi.ChartGroup SurfaceGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.WordApi.ChartGroup>(this, "SurfaceGroup", typeof(NetOffice.WordApi.ChartGroup));
            }
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff198336.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839101.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845234.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834948.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834564.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
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
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230485.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual NetOffice.WordApi.Enums.XlCategoryLabelLevel CategoryLabelLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlCategoryLabelLevel>(this, "CategoryLabelLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "CategoryLabelLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj232218.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual NetOffice.WordApi.Enums.XlSeriesNameLevel SeriesNameLevel
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.WordApi.Enums.XlSeriesNameLevel>(this, "SeriesNameLevel");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "SeriesNameLevel", value);
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get
        /// </summary>
        [SupportByVersion("Word", 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool HasHiddenContent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasHiddenContent");
            }
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj231924.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
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
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object SeriesCollection(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837270.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object SeriesCollection()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "SeriesCollection");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        /// <param name="separator">optional object separator</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", type, legendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194055.aspx </remarks>
        /// <param name="type">optional NetOffice.WordApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="legendKey">optional object legendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        /// <param name="showSeriesName">optional object showSeriesName</param>
        /// <param name="showCategoryName">optional object showCategoryName</param>
        /// <param name="showValue">optional object showValue</param>
        /// <param name="showPercentage">optional object showPercentage</param>
        /// <param name="showBubbleSize">optional object showBubbleSize</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyDataLabels(object type, object legendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyDataLabels", new object[] { type, legendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        /// <param name="typeName">optional object typeName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType, object typeName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType, typeName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyCustomType", chartType);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839889.aspx </remarks>
        /// <param name="x">Int32 x</param>
        /// <param name="y">Int32 y</param>
        /// <param name="elementID">Int32 elementID</param>
        /// <param name="arg1">Int32 arg1</param>
        /// <param name="arg2">Int32 arg2</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void GetChartElement(Int32 x, Int32 y, out Int32 elementID, out Int32 arg1, out Int32 arg2)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true, true, true);
            elementID = 0;
            arg1 = 0;
            arg2 = 0;
            object[] paramsArray = Invoker.ValidateParamsArray(x, y, elementID, arg1, arg2);
            Invoker.Method(this, "GetChartElement", paramsArray, modifiers);
            elementID = (Int32)paramsArray[2];
            arg1 = (Int32)paramsArray[3];
            arg2 = (Int32)paramsArray[4];
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
        /// <param name="source">string source</param>
        /// <param name="plotBy">optional object plotBy</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SetSourceData(string source, object plotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source, plotBy);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822921.aspx </remarks>
        /// <param name="source">string source</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SetSourceData(string source)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSourceData", source);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        /// <param name="type">optional object type</param>
        /// <param name="axisGroup">optional NetOffice.WordApi.Enums.XlAxisGroup AxisGroup = 1</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Axes(object type, object axisGroup)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type, axisGroup);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Axes()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193697.aspx </remarks>
        /// <param name="type">optional object type</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Axes(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Axes", type);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        /// <param name="format">optional object format</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void AutoFormat(Int32 gallery, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", gallery, format);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="gallery">Int32 gallery</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void AutoFormat(Int32 gallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoFormat", gallery);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197864.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SetBackgroundPicture(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetBackgroundPicture", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle, object extraTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle, extraTitle });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery, format);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", source, gallery, format, plotBy);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
        /// <param name="source">optional object source</param>
        /// <param name="gallery">optional object gallery</param>
        /// <param name="format">optional object format</param>
        /// <param name="plotBy">optional object plotBy</param>
        /// <param name="categoryLabels">optional object categoryLabels</param>
        /// <param name="seriesLabels">optional object seriesLabels</param>
        /// <param name="hasLegend">optional object hasLegend</param>
        /// <param name="title">optional object title</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196615.aspx </remarks>
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
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ChartWizard(object source, object gallery, object format, object plotBy, object categoryLabels, object seriesLabels, object hasLegend, object title, object categoryTitle, object valueTitle)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ChartWizard", new object[] { source, gallery, format, plotBy, categoryLabels, seriesLabels, hasLegend, title, categoryTitle, valueTitle });
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
        /// <param name="size">optional NetOffice.WordApi.Enums.XlPictureAppearance Size = 2</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format, object size)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format, size);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void CopyPicture()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void CopyPicture(object appearance)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823258.aspx </remarks>
        /// <param name="appearance">optional NetOffice.WordApi.Enums.XlPictureAppearance Appearance = 1</param>
        /// <param name="format">optional NetOffice.WordApi.Enums.XlCopyPictureFormat Format = -4147</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void CopyPicture(object appearance, object format)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "CopyPicture", appearance, format);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
        /// <param name="type">optional object type</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Paste(object type)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste", type);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839410.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Paste()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        /// <param name="interactive">optional object interactive</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual bool Export(string fileName, object filterName, object interactive)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName, filterName, interactive);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual bool Export(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195106.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="filterName">optional object filterName</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual bool Export(string fileName, object filterName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Export", fileName, filterName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839841.aspx </remarks>
        /// <param name="name">object name</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SetDefaultChart(object name)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultChart", name);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff845631.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyChartTemplate(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyChartTemplate", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839083.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SaveChartTemplate(string fileName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SaveChartTemplate", fileName);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193390.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ClearToMatchStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchStyle");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        /// <param name="chartType">optional object chartType</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout, object chartType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout, chartType);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840397.aspx </remarks>
        /// <param name="layout">Int32 layout</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void ApplyLayout(Int32 layout)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ApplyLayout", layout);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192135.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Refresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838552.aspx </remarks>
        /// <param name="element">NetOffice.OfficeApi.Enums.MsoChartElementType element</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void SetElement(NetOffice.OfficeApi.Enums.MsoChartElementType element)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetElement", element);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object AreaGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object AreaGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AreaGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object BarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object BarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "BarGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object ColumnGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object ColumnGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ColumnGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object LineGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object LineGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "LineGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object PieGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object PieGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "PieGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object DoughnutGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object DoughnutGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DoughnutGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object RadarGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object RadarGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "RadarGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object XYGroups(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups", index);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object XYGroups()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "XYGroups");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840074.aspx </remarks>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Delete()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        /// <param name="before">optional object before</param>
        /// <param name="after">optional object after</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Copy(object before, object after)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", before, after);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Copy()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192354.aspx </remarks>
        /// <param name="before">optional object before</param>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual void Copy(object before)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Copy", before);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
        /// <param name="replace">optional object replace</param>
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Select(object replace)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select", replace);
        }

        /// <summary>
        /// SupportByVersion Word 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191928.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 14, 15, 16)]
        public virtual object Select()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Word", 15, 16)]
        public virtual object FullSeriesCollection(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FullSeriesCollection", index);
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj229848.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Word", 15, 16)]
        public virtual object FullSeriesCollection()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "FullSeriesCollection");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Word", 15, 16)]
        public virtual void DeleteHiddenContent()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteHiddenContent");
        }

        /// <summary>
        /// SupportByVersion Word 15,16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/jj230203.aspx </remarks>
        [SupportByVersion("Word", 15, 16)]
        public virtual void ClearToMatchColorStyle()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearToMatchColorStyle");
        }

        #endregion

        #pragma warning restore
    }
}
