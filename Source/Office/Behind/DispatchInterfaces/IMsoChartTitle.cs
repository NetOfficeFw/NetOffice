using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// IMsoChartTitle
    /// </summary>
    [SyntaxBypass]
    public class IMsoChartTitle_ : COMObject, NetOffice.OfficeApi.IMsoChartTitle_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IMsoChartTitle_() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.OfficeApi.IMsoCharacters get_Characters(object start, object length)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoCharacters>(this, "Characters", typeof(NetOffice.OfficeApi.IMsoCharacters), start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        /// <param name="length">optional object length</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        public NetOffice.OfficeApi.IMsoCharacters Characters(object start, object length)
        {
            return get_Characters(start, length);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public NetOffice.OfficeApi.IMsoCharacters get_Characters(object start)
        {
            return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoCharacters>(this, "Characters", typeof(NetOffice.OfficeApi.IMsoCharacters), start);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Alias for get_Characters
        /// </summary>
        /// <param name="start">optional object start</param>
        [SupportByVersion("Office", 12, 14, 15, 16), Redirect("get_Characters")]
        public NetOffice.OfficeApi.IMsoCharacters Characters(object start)
        {
            return get_Characters(start);
        }

        #endregion

        #region Methods

        #endregion
    }

    /// <summary>
    /// DispatchInterface IMsoChartTitle 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class IMsoChartTitle : NetOffice.OfficeApi.Behind.IMsoChartTitle_, NetOffice.OfficeApi.IMsoChartTitle
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
                    _type = typeof(IMsoChartTitle);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not indented to use
        /// </summary>
        public IMsoChartTitle() : base()
        {

        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string Caption
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Caption");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Caption", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoCharacters Characters
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoCharacters>(this, "Characters", typeof(NetOffice.OfficeApi.IMsoCharacters));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ChartFont Font
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartFont>(this, "Font", typeof(NetOffice.OfficeApi.ChartFont));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object HorizontalAlignment
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "HorizontalAlignment");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "HorizontalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Double Left
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Left");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Left", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object Orientation
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "Orientation");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "Orientation", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public bool Shadow
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "Shadow");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Shadow", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string Text
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Text");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Text", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Double Top
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Top");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Top", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object VerticalAlignment
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "VerticalAlignment");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "VerticalAlignment", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Int32 ReadingOrder
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "ReadingOrder");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ReadingOrder", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object AutoScaleFont
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "AutoScaleFont");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "AutoScaleFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoInterior Interior
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.IMsoInterior>(this, "Interior", typeof(NetOffice.OfficeApi.IMsoInterior));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.ChartFillFormat Fill
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ChartFillFormat>(this, "Fill", typeof(NetOffice.OfficeApi.ChartFillFormat));
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoBorder Border
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
        public string Name
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Name");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16), ProxyResult]
        public object Parent
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
        public bool IncludeInLayout
        {
            get
            {
                return Factory.ExecuteBoolPropertyGet(this, "IncludeInLayout");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "IncludeInLayout", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.XlChartElementPosition Position
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.XlChartElementPosition>(this, "Position");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Position", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.IMsoChartFormat Format
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
        public Double Height
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Height");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16), ProxyResult]
        public object Application
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
        public Int32 Creator
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
        public Double Width
        {
            get
            {
                return Factory.ExecuteDoublePropertyGet(this, "Width");
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public string Formula
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Formula");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Formula", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public string FormulaR1C1
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "FormulaR1C1");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FormulaR1C1", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public string FormulaLocal
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "FormulaLocal");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FormulaLocal", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public string FormulaR1C1Local
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "FormulaR1C1Local");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FormulaR1C1Local", value);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object Delete()
        {
            return Factory.ExecuteVariantMethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public object Select()
        {
            return Factory.ExecuteVariantMethodGet(this, "Select");
        }

        #endregion

        #pragma warning restore
    }
}
