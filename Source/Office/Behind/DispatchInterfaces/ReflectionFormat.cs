using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ReflectionFormat 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863140.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ReflectionFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.ReflectionFormat
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
                    _type = typeof(ReflectionFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ReflectionFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861491.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoReflectionType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoReflectionType>(this, "Type");
            }
            set
            {
                Factory.ExecuteEnumPropertySet(this, "Type", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861198.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Transparency
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
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862407.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Size
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Size");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Size", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864080.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Offset
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Offset");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Offset", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865483.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Blur
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Blur");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Blur", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
