using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface LineFormat 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.ExcelApi.LineFormat")]
    public class LineFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.LineFormat
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
                    _type = typeof(LineFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public LineFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat BackColor
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "BackColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "BackColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadLength BeginArrowheadLength
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadLength>(this, "BeginArrowheadLength");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BeginArrowheadLength", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadStyle BeginArrowheadStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadStyle>(this, "BeginArrowheadStyle");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BeginArrowheadStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadWidth BeginArrowheadWidth
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadWidth>(this, "BeginArrowheadWidth");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "BeginArrowheadWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLineDashStyle DashStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLineDashStyle>(this, "DashStyle");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "DashStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadLength EndArrowheadLength
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadLength>(this, "EndArrowheadLength");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "EndArrowheadLength", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadStyle EndArrowheadStyle
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadStyle>(this, "EndArrowheadStyle");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "EndArrowheadStyle", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoArrowheadWidth EndArrowheadWidth
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoArrowheadWidth>(this, "EndArrowheadWidth");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "EndArrowheadWidth", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.ColorFormat ForeColor
        {
            get
            {
                return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.ColorFormat>(this, "ForeColor", typeof(NetOffice.OfficeApi.ColorFormat));
            }
            set
            {
                Factory.ExecuteReferencePropertySet(this, "ForeColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoPatternType Pattern
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoPatternType>(this, "Pattern");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Pattern", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoLineStyle Style
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoLineStyle>(this, "Style");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Style", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Transparency
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Transparency");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Transparency", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState Visible
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "Visible");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Visible", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Single Weight
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Weight");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Weight", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi.Enums.MsoTriState InsetPen
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoTriState>(this, "InsetPen");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "InsetPen", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
