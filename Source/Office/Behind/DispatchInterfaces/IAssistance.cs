using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface IAssistance 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864589.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class IAssistance : COMObject, NetOffice.OfficeApi.IAssistance
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
                    _contractType = typeof(NetOffice.OfficeApi.IAssistance);
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
                    _type = typeof(IAssistance);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IAssistance() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        /// <param name="helpId">optional string HelpId = </param>
        /// <param name="scope">optional string Scope = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowHelp(object helpId, object scope)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp", helpId, scope);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowHelp()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp");
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860570.aspx </remarks>
        /// <param name="helpId">optional string HelpId = </param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ShowHelp(object helpId)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ShowHelp", helpId);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862805.aspx </remarks>
        /// <param name="query">string query</param>
        /// <param name="scope">optional string Scope = </param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SearchHelp(string query, object scope)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SearchHelp", query, scope);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862805.aspx </remarks>
        /// <param name="query">string query</param>
        [CustomMethod]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SearchHelp(string query)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SearchHelp", query);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861230.aspx </remarks>
        /// <param name="helpId">string helpId</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void SetDefaultContext(string helpId)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDefaultContext", helpId);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865260.aspx </remarks>
        /// <param name="helpId">string helpId</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void ClearDefaultContext(string helpId)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ClearDefaultContext", helpId);
        }

        #endregion

        #pragma warning restore
    }
}
