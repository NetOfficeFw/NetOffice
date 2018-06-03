using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface WorkflowTemplate 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863138.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class WorkflowTemplate : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.WorkflowTemplate
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
                    _type = typeof(WorkflowTemplate);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public WorkflowTemplate() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861378.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string Id
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Id");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861417.aspx </remarks>
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
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863121.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string Description
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "Description");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861722.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string DocumentLibraryName
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "DocumentLibraryName");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860562.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public string DocumentLibraryURL
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "DocumentLibraryURL");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863678.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public Int32 Show()
        {
            return Factory.ExecuteInt32MethodGet(this, "Show");
        }

        #endregion

        #pragma warning restore
    }
}
