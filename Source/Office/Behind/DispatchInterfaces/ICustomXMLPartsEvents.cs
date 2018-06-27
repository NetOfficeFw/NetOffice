using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// DispatchInterface ICustomXMLPartsEvents 
    /// SupportByVersion Office, 12,14,15,16
    /// </summary>
    [SupportByVersion("Office", 12, 14, 15, 16)]
    [EntityType(EntityType.IsDispatchInterface)]
    public class ICustomXMLPartsEvents : COMObject, NetOffice.OfficeApi.ICustomXMLPartsEvents
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
                    _contractType = typeof(NetOffice.OfficeApi.ICustomXMLPartsEvents);
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
                    _type = typeof(ICustomXMLPartsEvents);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public ICustomXMLPartsEvents() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="newPart">NetOffice.OfficeApi.CustomXMLPart newPart</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PartAfterAdd(NetOffice.OfficeApi.CustomXMLPart newPart)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PartAfterAdd", newPart);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="oldPart">NetOffice.OfficeApi.CustomXMLPart oldPart</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PartBeforeDelete(NetOffice.OfficeApi.CustomXMLPart oldPart)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PartBeforeDelete", oldPart);
        }

        /// <summary>
        /// SupportByVersion Office 12, 14, 15, 16
        /// </summary>
        /// <param name="part">NetOffice.OfficeApi.CustomXMLPart part</param>
        [SupportByVersion("Office", 12, 14, 15, 16)]
        public virtual void PartAfterLoad(NetOffice.OfficeApi.CustomXMLPart part)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "PartAfterLoad", part);
        }

        #endregion

        #pragma warning restore
    }
}
