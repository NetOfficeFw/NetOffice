using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IEnumPrivacyRecords 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IEnumPrivacyRecords : COMObject, NetOffice.MSHTMLApi.IEnumPrivacyRecords
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IEnumPrivacyRecords);
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
                    _type = typeof(IEnumPrivacyRecords);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IEnumPrivacyRecords() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 reset()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "reset");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSize">Int32 pSize</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetSize(out Int32 pSize)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pSize = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pSize);
			object returnItem = Invoker.MethodReturn(this, "GetSize", paramsArray, modifiers);
			pSize = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pState">Int32 pState</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetPrivacyImpacted(out Int32 pState)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pState = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pState);
			object returnItem = Invoker.MethodReturn(this, "GetPrivacyImpacted", paramsArray, modifiers);
			pState = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pbstrUrl">string pbstrUrl</param>
		/// <param name="pbstrPolicyRef">string pbstrPolicyRef</param>
		/// <param name="pdwReserved">Int32 pdwReserved</param>
		/// <param name="pdwPrivacyFlags">Int32 pdwPrivacyFlags</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 Next(out string pbstrUrl, out string pbstrPolicyRef, out Int32 pdwReserved, out Int32 pdwPrivacyFlags)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pbstrUrl = string.Empty;
			pbstrPolicyRef = string.Empty;
			pdwReserved = 0;
			pdwPrivacyFlags = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstrUrl, pbstrPolicyRef, pdwReserved, pdwPrivacyFlags);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray, modifiers);
			pbstrUrl = paramsArray[0] as string;
			pbstrPolicyRef = paramsArray[1] as string;
			pdwReserved = (Int32)paramsArray[2];
			pdwPrivacyFlags = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

