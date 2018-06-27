using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OfficeApi;

namespace NetOffice.OfficeApi.Behind
{
    /// <summary>
    /// Interface IConverterPreferences 
    /// SupportByVersion Office, 14,15,16
    /// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864179.aspx </remarks>
    [SupportByVersion("Office", 14, 15, 16)]
    [EntityType(EntityType.IsInterface)]
    public class IConverterPreferences : COMObject, NetOffice.OfficeApi.IConverterPreferences
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
                    _contractType = typeof(NetOffice.OfficeApi.IConverterPreferences);
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
                    _type = typeof(IConverterPreferences);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IConverterPreferences() : base()
		{

		}

		#endregion

        #region Properties

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863141.aspx </remarks>
        /// <param name="pfMacroEnabled">Int32 pfMacroEnabled</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrGetMacroEnabled(out Int32 pfMacroEnabled)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            pfMacroEnabled = 0;
            object[] paramsArray = new object[] { pfMacroEnabled };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrGetMacroEnabled", paramsArray, modifiers);

            pfMacroEnabled = (Int32)paramsArray[0];
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865570.aspx </remarks>
        /// <param name="pFormat">Int32 pFormat</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrCheckFormat(out Int32 pFormat)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            pFormat = 0;
            object[] paramsArray = new object[] { pFormat };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrCheckFormat", paramsArray, modifiers);

            pFormat = (Int32)paramsArray[0];
            return returnItem;
        }

        /// <summary>
        /// SupportByVersion Office 14, 15, 16
        /// </summary>
        /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860851.aspx </remarks>
        /// <param name="pfLossySave">Int32 pfLossySave</param>
        [SupportByVersion("Office", 14, 15, 16)]
        public virtual Int32 HrGetLossySave(out Int32 pfLossySave)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
            pfLossySave = 0;
            object[] paramsArray = new object[] { pfLossySave };

            Int32 returnItem = InvokerService.InvokeInternal.ExecuteInt32MethodGetExtended(this, "HrGetLossySave", paramsArray, modifiers);

            pfLossySave = (Int32)paramsArray[0];
            return returnItem;
        }

        #endregion

        #pragma warning restore
    }
}
