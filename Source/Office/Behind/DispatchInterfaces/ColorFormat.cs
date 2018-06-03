using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ColorFormat 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    [Duplicate("NetOffice.ExcelApi.ColorFormat")]
    public class ColorFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ColorFormat
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
                    _type = typeof(ColorFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ColorFormat() : base()
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
        public object Parent
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
        public Int32 RGB
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "RGB");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "RGB", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Int32 SchemeColor
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "SchemeColor");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "SchemeColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoColorType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoColorType>(this, "Type");
            }
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public Single TintAndShade
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "TintAndShade");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "TintAndShade", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoThemeColorIndex ObjectThemeColor
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoThemeColorIndex>(this, "ObjectThemeColor");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "ObjectThemeColor", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Brightness
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Brightness");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Brightness", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
