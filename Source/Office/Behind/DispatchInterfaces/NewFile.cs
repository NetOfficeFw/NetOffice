using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface NewFile 
    /// SupportByVersion Office, 10,11,12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862417.aspx </remarks>
    [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class NewFile : NetOffice.OfficeApi.Behind._IMsoDispObj, NetOffice.OfficeApi.NewFile
    {
        #pragma warning disable

        #region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.OfficeApi.NewFile);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(NewFile);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public NewFile() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        /// <param name="action">optional object action</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Add(string fileName, object section, object displayName, object action)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Add", fileName, section, displayName, action);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Add(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Add", fileName);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Add(string fileName, object section)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Add", fileName, section);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860279.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Add(string fileName, object section, object displayName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Add", fileName, section, displayName);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        /// <param name="action">optional object action</param>
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Remove(string fileName, object section, object displayName, object action)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Remove", fileName, section, displayName, action);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Remove(string fileName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Remove", fileName);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Remove(string fileName, object section)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Remove", fileName, section);
        }

        /// <summary>
        /// SupportByVersion Office 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860573.aspx </remarks>
        /// <param name="fileName">string fileName</param>
        /// <param name="section">optional object section</param>
        /// <param name="displayName">optional object displayName</param>
        [CustomMethod]
        [SupportByVersion("Office", 10, 11, 12, 14, 15, 16)]
        public virtual bool Remove(string fileName, object section, object displayName)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "Remove", fileName, section, displayName);
        }

        #endregion

        #pragma warning restore
    }
}
