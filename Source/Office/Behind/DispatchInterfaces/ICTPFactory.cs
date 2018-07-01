using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ICTPFactory
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864938.aspx </remarks>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ICTPFactory : COMObject, NetOffice.OfficeApi.ICTPFactory
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
                    _contractType = typeof(NetOffice.OfficeApi.ICTPFactory);
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
                    _type = typeof(ICTPFactory);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICTPFactory() : base()
		{

		}



        public ICTPFactory(Core factory, ICOMObject caller, object comProxy) : base(factory, caller, comProxy)
        {

        }
        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx </remarks>
        /// <param name="cTPAxID">string cTPAxID</param>
        /// <param name="cTPTitle">string cTPTitle</param>
        /// <param name="cTPParentWindow">optional object cTPParentWindow</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        [BaseResult]
        public virtual NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle, object cTPParentWindow)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi._CustomTaskPane>(this, "CreateCTP", cTPAxID, cTPTitle, cTPParentWindow);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860563.aspx </remarks>
        /// <param name="cTPAxID">string cTPAxID</param>
        /// <param name="cTPTitle">string cTPTitle</param>
        [CustomMethod]
        [BaseResult]
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual NetOffice.OfficeApi._CustomTaskPane CreateCTP(string cTPAxID, string cTPTitle)
        {
            return InvokerService.InvokeInternal.ExecuteBaseReferenceMethodGet<NetOffice.OfficeApi._CustomTaskPane>(this, "CreateCTP", cTPAxID, cTPTitle);
        }

        #endregion

        #pragma warning restore
    }
}
