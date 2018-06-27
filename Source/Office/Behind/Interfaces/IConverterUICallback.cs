using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface IConverterUICallback 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863370.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IConverterUICallback : COMObject, NetOffice.OfficeApi.IConverterUICallback
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
                    _contractType = typeof(NetOffice.OfficeApi.IConverterUICallback);
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
                    _type = typeof(IConverterUICallback);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IConverterUICallback() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861826.aspx </remarks>
        /// <param name="uPercentComplete">UIntPtr uPercentComplete</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrReportProgress(UIntPtr uPercentComplete)
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "HrReportProgress", uPercentComplete);
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861376.aspx </remarks>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrCaption">string bstrCaption</param>
        /// <param name="uType">UIntPtr uType</param>
        /// <param name="pidResult">Int32 pidResult</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrMessageBox(string bstrText, string bstrCaption, UIntPtr uType, out Int32 pidResult)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, true);
            pidResult = 0;
            object[] paramsArray = new object[] { bstrText, bstrCaption, uType, pidResult };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrMessageBox", paramsArray, modifiers);

            pidResult = (Int32)paramsArray[3];
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861803.aspx </remarks>
        /// <param name="bstrText">string bstrText</param>
        /// <param name="bstrCaption">string bstrCaption</param>
        /// <param name="pbstrInput">string pbstrInput</param>
        /// <param name="fPassword">Int32 fPassword</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrInputBox(string bstrText, string bstrCaption, out string pbstrInput, Int32 fPassword)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true, false);
            pbstrInput = string.Empty;
            object[] paramsArray = new object[] { bstrText, bstrCaption, pbstrInput, fPassword };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrInputBox", paramsArray, modifiers);

            pbstrInput = paramsArray[2] as string;
            return returnItem;
        }

        #endregion

        #pragma warning restore
    }
}
