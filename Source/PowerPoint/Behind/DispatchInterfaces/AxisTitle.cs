using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.PowerPointApi;

namespace NetOffice.PowerPointApi.Behind
{
    /// <summary>
    /// AxisTitle
    /// </summary>
    [SyntaxBypass]
    public class AxisTitle_ : COMObject,  NetOffice.PowerPointApi.AxisTitle_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public AxisTitle_() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartCharacters get_Characters(object start, object length)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartCharacters>(this, "Characters", typeof(NetOffice.PowerPointApi.ChartCharacters), start, length);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.PowerPointApi.ChartCharacters Characters(object start, object length)
        {
            return get_Characters(start, length);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartCharacters get_Characters(object start)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartCharacters>(this, "Characters", typeof(NetOffice.PowerPointApi.ChartCharacters), start);
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
        /// <param name="start">optional object start</param>
        [SupportByVersion("PowerPoint", 14, 15, 16), Redirect("get_Characters")]
        public virtual NetOffice.PowerPointApi.ChartCharacters Characters(object start)
        {
            return get_Characters(start);
        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface AxisTitle
    /// SupportByVersion PowerPoint, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745597.aspx </remarks>
    public class AxisTitle : AxisTitle_, NetOffice.PowerPointApi.AxisTitle
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
                    _contractType = typeof(NetOffice.PowerPointApi.AxisTitle);
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
                    _type = typeof(AxisTitle);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public AxisTitle() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746628.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Caption
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Caption");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Caption", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745546.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.ChartCharacters Characters
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartCharacters>(this, "Characters", typeof(NetOffice.PowerPointApi.ChartCharacters));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartFont Font
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFont>(this, "Font", typeof(NetOffice.PowerPointApi.ChartFont));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746675.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object HorizontalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744140.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Double Left
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Left");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746033.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Orientation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Orientation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Orientation", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744568.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual bool Shadow
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Shadow");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Shadow", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746196.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Text
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744545.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Double Top
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Top");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743882.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object VerticalAlignment
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "VerticalAlignment");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object AutoScaleFont
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "AutoScaleFont");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "AutoScaleFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.Interior Interior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.Interior>(this, "Interior", typeof(NetOffice.PowerPointApi.Interior));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartFillFormat Fill
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartFillFormat>(this, "Fill", typeof(NetOffice.PowerPointApi.ChartFillFormat));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.PowerPointApi.ChartBorder Border
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.PowerPointApi.ChartBorder>(this, "Border", typeof(NetOffice.PowerPointApi.ChartBorder));
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746692.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Name
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745911.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff743946.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual bool IncludeInLayout
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IncludeInLayout");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IncludeInLayout", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745347.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual NetOffice.PowerPointApi.Enums.XlChartElementPosition Position
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.PowerPointApi.Enums.XlChartElementPosition>(this, "Position");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "Position", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745870.aspx </remarks>
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
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746782.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746190.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744796.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Int32 ReadingOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "ReadingOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ReadingOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff746563.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Double Height
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745147.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual Double Width
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745933.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string Formula
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Formula");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Formula", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744316.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string FormulaR1C1
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaR1C1");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaR1C1", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff745627.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string FormulaLocal
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaLocal");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaLocal", value);
            }
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744715.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual string FormulaR1C1Local
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FormulaR1C1Local");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "FormulaR1C1Local", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744056.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Delete()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion PowerPoint 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff744775.aspx </remarks>
        [SupportByVersion("PowerPoint", 14, 15, 16)]
        public virtual object Select()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
        }

        #endregion

        #pragma warning restore
    }
}
