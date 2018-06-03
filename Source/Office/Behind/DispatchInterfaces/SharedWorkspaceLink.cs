using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface SharedWorkspaceLink 
    /// SupportByVersion Office, 11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865254.aspx </remarks>
    [SupportByVersion("Office", 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class SharedWorkspaceLink : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.SharedWorkspaceLink
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
                    _type = typeof(SharedWorkspaceLink);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public SharedWorkspaceLink() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863119.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string URL
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "URL");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "URL", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860499.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string Description
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Description");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Description", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get/Set
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861886.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string Notes
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Notes");
            }
            set
            {
                Factory.ExecuteValuePropertySet(this, "Notes", value);
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863662.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string CreatedBy
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "CreatedBy");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861532.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public object CreatedDate
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "CreatedDate");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861094.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public string ModifiedBy
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ModifiedBy");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860294.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public object ModifiedDate
        {
            get
            {
                return Factory.ExecuteVariantPropertyGet(this, "ModifiedDate");
            }
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863516.aspx </remarks>
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862046.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public void Save()
        {
            Factory.ExecuteMethod(this, "Save");
        }

        /// <summary>
        /// SupportByVersion Office 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863052.aspx </remarks>
        [SupportByVersion("Office", 11, 12, 14, 15, 16)]
        public void Delete()
        {
            Factory.ExecuteMethod(this, "Delete");
        }

        #endregion

        #pragma warning restore
    }
}
