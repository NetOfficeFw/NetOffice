using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface UserPermission 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860810.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class UserPermission : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.UserPermission
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
                    _type = typeof(UserPermission);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public UserPermission() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862102.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string UserId
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "UserId");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862094.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public Int32 Permission
        {
            get
            {
                return Factory.ExecuteInt32PropertyGet(this, "Permission");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Permission", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862529.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public object ExpirationDate
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ExpirationDate");
            }
            set
            {
                Factory.ExecuteVariantPropertySet(this, "ExpirationDate", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861797.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16), ProxyResult]
        public object Parent
        {
            get
            {
                return Factory.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864865.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public void Remove()
        {
            Factory.ExecuteMethod(this, "Remove");
        }

        #endregion

        #pragma warning restore
    }
}
