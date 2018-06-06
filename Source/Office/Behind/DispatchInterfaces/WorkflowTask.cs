using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface WorkflowTask 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863345.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class WorkflowTask : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.WorkflowTask
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
                    _type = typeof(WorkflowTask);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public WorkflowTask() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863854.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Id
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860221.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string ListID
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "ListID");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861217.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string WorkflowID
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "WorkflowID");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865575.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Name
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865524.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string Description
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
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864072.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string AssignedTo
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "AssignedTo");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861453.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual string CreatedBy
        {
            get
            {
                return Factory.ExecuteStringPropertyGet(this, "CreatedBy");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862844.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual DateTime DueDate
        {
            get
            {
                return Factory.ExecuteDateTimePropertyGet(this, "DueDate");
            }
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861139.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual DateTime CreatedDate
        {
            get
            {
                return Factory.ExecuteDateTimePropertyGet(this, "CreatedDate");
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863536.aspx </remarks>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual Int32 Show()
        {
            return Factory.ExecuteInt32MethodGet(this, "Show");
        }

        #endregion

        #pragma warning restore
    }
}
