using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.MSHTMLApi;

namespace NetOffice.MSHTMLApi.Behind
{
	/// <summary>
	/// Interface IWPCBlockedUrls 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IWPCBlockedUrls : COMObject, NetOffice.MSHTMLApi.IWPCBlockedUrls
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
                    _contractType = typeof(NetOffice.MSHTMLApi.IWPCBlockedUrls);
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
                    _type = typeof(IWPCBlockedUrls);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IWPCBlockedUrls() : base()
		{

		}

		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pdwCount">Int32 pdwCount</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetCount(out Int32 pdwCount)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pdwCount = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pdwCount);
			object returnItem = Invoker.MethodReturn(this, "GetCount", paramsArray, modifiers);
			pdwCount = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwIdx">Int32 dwIdx</param>
		/// <param name="pbstrUrl">string pbstrUrl</param>
		[SupportByVersion("MSHTML", 4)]
		public virtual Int32 GetUrl(Int32 dwIdx, out string pbstrUrl)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pbstrUrl = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(dwIdx, pbstrUrl);
			object returnItem = Invoker.MethodReturn(this, "GetUrl", paramsArray, modifiers);
			pbstrUrl = paramsArray[1] as string;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}

