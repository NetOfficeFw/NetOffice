using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface WebPageFont 
    /// SupportByVersion Office, 9,10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864941.aspx </remarks>
    [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class WebPageFont : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.WebPageFont
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
                    _type = typeof(WebPageFont);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public WebPageFont() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865546.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string ProportionalFont
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ProportionalFont");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProportionalFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863960.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single ProportionalFontSize
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "ProportionalFontSize");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "ProportionalFontSize", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865471.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public string FixedWidthFont
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "FixedWidthFont");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FixedWidthFont", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863486.aspx </remarks>
        [SupportByVersion("Office", 9, 10, 11, 12, 14, 15, 16)]
        public Single FixedWidthFontSize
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "FixedWidthFontSize");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "FixedWidthFontSize", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
