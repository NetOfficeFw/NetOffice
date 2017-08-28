using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSHTMLApi
{
	/// <summary>
	/// Interface IMarkupContainer2 
	/// SupportByVersion MSHTML, 4
	/// </summary>
	[SupportByVersion("MSHTML", 4)]
	[EntityType(EntityType.IsInterface)]
 	public class IMarkupContainer2 : IMarkupContainer
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
                    _type = typeof(IMarkupContainer2);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMarkupContainer2(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMarkupContainer2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMarkupContainer2(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="ppChangeLog">NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog</param>
		/// <param name="fForward">Int32 fForward</param>
		/// <param name="fBackward">Int32 fBackward</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 CreateChangeLog(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog, Int32 fForward, Int32 fBackward)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);
			ppChangeLog = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, ppChangeLog, fForward, fBackward);
			object returnItem = Invoker.MethodReturn(this, "CreateChangeLog", paramsArray, modifiers);
            if (paramsArray[1] is MarshalByRefObject)
                ppChangeLog = new NetOffice.MSHTMLApi.IHTMLChangeLog(this, paramsArray[1]);
            else
                ppChangeLog = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="pdwCookie">Int32 pdwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 RegisterForDirtyRange(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out Int32 pdwCookie)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pdwCookie = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, pdwCookie);
			object returnItem = Invoker.MethodReturn(this, "RegisterForDirtyRange", paramsArray, modifiers);
			pdwCookie = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 UnRegisterForDirtyRange(Int32 dwCookie)
		{
			return Factory.ExecuteInt32MethodGet(this, "UnRegisterForDirtyRange", dwCookie);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		/// <param name="pIPointerBegin">NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin</param>
		/// <param name="pIPointerEnd">NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetAndClearDirtyRange(Int32 dwCookie, NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin, NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetAndClearDirtyRange", dwCookie, pIPointerBegin, pIPointerEnd);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetVersionNumber()
		{
			return Factory.ExecuteInt32MethodGet(this, "GetVersionNumber");
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// </summary>
		/// <param name="ppElementMaster">NetOffice.MSHTMLApi.IHTMLElement ppElementMaster</param>
		[SupportByVersion("MSHTML", 4)]
		public Int32 GetMasterElement(out NetOffice.MSHTMLApi.IHTMLElement ppElementMaster)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppElementMaster = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppElementMaster);
			object returnItem = Invoker.MethodReturn(this, "GetMasterElement", paramsArray, modifiers);
            if (paramsArray[0] is MarshalByRefObject)
                ppElementMaster = new NetOffice.MSHTMLApi.IHTMLElement(this, paramsArray[0]);
            else
                ppElementMaster = null;
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion

		#pragma warning restore
	}
}
