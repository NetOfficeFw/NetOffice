using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// Interface IEnumPrivacyRecords 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IEnumPrivacyRecords : COMObject
	{
		#pragma warning disable
		#region Type Information

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
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IEnumPrivacyRecords(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumPrivacyRecords(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 reset()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "reset", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pSize">Int32 pSize</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetSize(out Int32 pSize)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pSize = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pSize);
			object returnItem = Invoker.MethodReturn(this, "GetSize", paramsArray);
			pSize = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pState">Int32 pState</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetPrivacyImpacted(out Int32 pState)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pState = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pState);
			object returnItem = Invoker.MethodReturn(this, "GetPrivacyImpacted", paramsArray);
			pState = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pbstrUrl">string pbstrUrl</param>
		/// <param name="pbstrPolicyRef">string pbstrPolicyRef</param>
		/// <param name="pdwReserved">Int32 pdwReserved</param>
		/// <param name="pdwPrivacyFlags">Int32 pdwPrivacyFlags</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Next(out string pbstrUrl, out string pbstrPolicyRef, out Int32 pdwReserved, out Int32 pdwPrivacyFlags)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			pbstrUrl = string.Empty;
			pbstrPolicyRef = string.Empty;
			pdwReserved = 0;
			pdwPrivacyFlags = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pbstrUrl, pbstrPolicyRef, pdwReserved, pdwPrivacyFlags);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray);
			pbstrUrl = (string)paramsArray[0];
			pbstrPolicyRef = (string)paramsArray[1];
			pdwReserved = (Int32)paramsArray[2];
			pdwPrivacyFlags = (Int32)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}