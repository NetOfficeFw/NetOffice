using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IEnumRegisterWordW 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IEnumRegisterWordW : COMObject
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
                    _type = typeof(IEnumRegisterWordW);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IEnumRegisterWordW(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IEnumRegisterWordW(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IEnumRegisterWordW(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppEnum">NetOffice.MSHTMLApi.IEnumRegisterWordW ppEnum</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 Clone(out NetOffice.MSHTMLApi.IEnumRegisterWordW ppEnum)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppEnum = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppEnum);
			object returnItem = Invoker.MethodReturn(this, "Clone", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppEnum = new NetOffice.MSHTMLApi.IEnumRegisterWordW(this, paramsArray[0]);
            else
                ppEnum = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ulCount">Int32 ulCount</param>
		/// <param name="rgRegisterWord">__MIDL___MIDL_itf_mshtml_0001_0042_0002 rgRegisterWord</param>
		/// <param name="pcFetched">Int32 pcFetched</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 Next(Int32 ulCount, out __MIDL___MIDL_itf_mshtml_0001_0042_0002 rgRegisterWord, out Int32 pcFetched)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true);
			rgRegisterWord = new NetOffice.MSHTMLApi.__MIDL___MIDL_itf_mshtml_0001_0042_0002();
			pcFetched = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(ulCount, rgRegisterWord, pcFetched);
			object returnItem = Invoker.MethodReturn(this, "Next", paramsArray, modifiers);
			rgRegisterWord = (__MIDL___MIDL_itf_mshtml_0001_0042_0002)paramsArray[1];
			pcFetched = (Int32)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

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
		/// <param name="ulCount">Int32 ulCount</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 Skip(Int32 ulCount)
		{
			return Factory.ExecuteInt32MethodGet(this, "Skip", ulCount);
		}

		#endregion

		#pragma warning restore
	}
}
