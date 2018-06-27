using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface IMsoSeries 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IMsoSeries : COMObject, NetOffice.OfficeApi.IMsoSeries
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
                    _contractType = typeof(NetOffice.OfficeApi.IMsoSeries);
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
                    _type = typeof(IMsoSeries);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IMsoSeries() : base()
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
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlAxisGroup AxisGroup
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlAxisGroup>(this, "AxisGroup");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "AxisGroup", value);
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
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoBorder>(this, "Border", typeof(NetOffice.OfficeApi.IMsoBorder));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoErrorBars ErrorBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoErrorBars>(this, "ErrorBars", typeof(NetOffice.OfficeApi.IMsoErrorBars));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Explosion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Explosion");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Explosion", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasDataLabels
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasDataLabels");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasDataLabels", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasErrorBars
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasErrorBars");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasErrorBars", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoInterior Interior
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoInterior>(this, "Interior", typeof(NetOffice.OfficeApi.IMsoInterior));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ChartFillFormat Fill
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartFillFormat>(this, "Fill", typeof(NetOffice.OfficeApi.ChartFillFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool InvertIfNegative
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "InvertIfNegative");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvertIfNegative", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 MarkerBackgroundColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerBackgroundColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerBackgroundColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlColorIndex MarkerBackgroundColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "MarkerBackgroundColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerBackgroundColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 MarkerForegroundColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerForegroundColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerForegroundColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlColorIndex MarkerForegroundColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "MarkerForegroundColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerForegroundColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 MarkerSize
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MarkerSize");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MarkerSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlMarkerStyle MarkerStyle
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlMarkerStyle>(this, "MarkerStyle");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "MarkerStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlChartPictureType PictureType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartPictureType>(this, "PictureType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "PictureType", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Double PictureUnit
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PictureUnit");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureUnit", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 PlotOrder
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PlotOrder");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PlotOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Smooth
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Smooth");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Smooth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Values
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "Values");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "Values", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object XValues
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "XValues");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "XValues", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object BubbleSizes
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteVariantPropertyGet(this, "BubbleSizes");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteVariantPropertySet(this, "BubbleSizes", value);
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
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlBarShape>(this, "BarShape");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "BarShape", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ApplyPictToSides
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToSides");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToSides", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ApplyPictToFront
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToFront");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToFront", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool ApplyPictToEnd
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "ApplyPictToEnd");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ApplyPictToEnd", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool Has3DEffect
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Has3DEffect");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "Has3DEffect", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
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
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual bool HasLeaderLines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "HasLeaderLines");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "HasLeaderLines", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.IMsoLeaderLines LeaderLines
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoLeaderLines>(this, "LeaderLines", typeof(NetOffice.OfficeApi.IMsoLeaderLines));
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
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoChartFormat>(this, "Format", typeof(NetOffice.OfficeApi.IMsoChartFormat));
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
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Application");
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
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Double PictureUnit2
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteDoublePropertyGet(this, "PictureUnit2");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "PictureUnit2", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 PlotColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "PlotColorIndex");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 InvertColor
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "InvertColor");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "InvertColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.XlColorIndex InvertColorIndex
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlColorIndex>(this, "InvertColorIndex");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "InvertColorIndex", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 15,16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 15, 16)]
        public virtual bool IsFiltered
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsFiltered");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsFiltered", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        /// <param name="hasLeaderLines">optional object hasLeaderLines</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object _ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object _ApplyDataLabels()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object _ApplyDataLabels(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object _ApplyDataLabels(object type, object iMsoLegendKey)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey);
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
        public virtual object _ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "_ApplyDataLabels", type, iMsoLegendKey, autoText);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ClearFormats()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ClearFormats");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Copy()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Copy");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object DataLabels(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataLabels", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object DataLabels()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "DataLabels");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Delete()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
        /// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
        /// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
        /// <param name="amount">optional object amount</param>
        /// <param name="minusValues">optional object minusValues</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type, object amount, object minusValues)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", new object[] { direction, include, type, amount, minusValues });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
        /// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
        /// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="direction">NetOffice.OfficeApi.Enums.XlErrorBarDirection direction</param>
        /// <param name="include">NetOffice.OfficeApi.Enums.XlErrorBarInclude include</param>
        /// <param name="type">NetOffice.OfficeApi.Enums.XlErrorBarType type</param>
        /// <param name="amount">optional object amount</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ErrorBar(NetOffice.OfficeApi.Enums.XlErrorBarDirection direction, NetOffice.OfficeApi.Enums.XlErrorBarInclude include, NetOffice.OfficeApi.Enums.XlErrorBarType type, object amount)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ErrorBar", direction, include, type, amount);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Paste()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Paste");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Points(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Points", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Points()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Points");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Select()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Select");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="index">optional object index</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Trendlines(object index)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trendlines", index);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object Trendlines()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Trendlines");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="chartType">NetOffice.OfficeApi.Enums.XlChartType chartType</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 ApplyCustomType(NetOffice.OfficeApi.Enums.XlChartType chartType)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ApplyCustomType", chartType);
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize, object separator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize, separator });
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ApplyDataLabels()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ApplyDataLabels(object type)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="type">optional NetOffice.OfficeApi.Enums.XlDataLabelsType Type = 2</param>
        /// <param name="iMsoLegendKey">optional object iMsoLegendKey</param>
        /// <param name="autoText">optional object autoText</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey, autoText);
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", type, iMsoLegendKey, autoText, hasLeaderLines);
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName });
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName });
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue });
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage });
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
        public virtual object ApplyDataLabels(object type, object iMsoLegendKey, object autoText, object hasLeaderLines, object showSeriesName, object showCategoryName, object showValue, object showPercentage, object showBubbleSize)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "ApplyDataLabels", new object[] { type, iMsoLegendKey, autoText, hasLeaderLines, showSeriesName, showCategoryName, showValue, showPercentage, showBubbleSize });
        }

        #endregion

        #pragma warning restore
    }
}
