using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumPrivacyRecords 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IEnumPrivacyRecords : COMObject
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
                    _type = typeof(IEnumPrivacyRecords);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IEnumPrivacyRecords(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 reset()
		{
			return Factory.ExecuteInt32MethodGet(this, "reset");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pSize">Int32 pSize</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetSize(out Int32 pSize)
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
		public Int32 GetPrivacyImpacted(out Int32 pState)
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
		public Int32 Next(out string pbstrUrl, out string pbstrPolicyRef, out Int32 pdwReserved, out Int32 pdwPrivacyFlags)
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
