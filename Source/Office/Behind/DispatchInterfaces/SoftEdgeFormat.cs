using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SoftEdgeFormat 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863361.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class SoftEdgeFormat : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SoftEdgeFormat
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
                    _type = typeof(SoftEdgeFormat);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SoftEdgeFormat() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865253.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public NetOffice.OfficeApi.Enums.MsoSoftEdgeType Type
        {
            get
            {
                return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSoftEdgeType>(this, "Type");
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862536.aspx </remarks>
        [SupportByVersion("Office", 14, 15, 16)]
        public Single Radius
        {
            get
            {
                return Factory.ExecuteSinglePropertyGet(this, "Radius");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Radius", value);
            }
        }

        #endregion

        #region Methods

        #endregion

        #pragma warning restore
    }
}
