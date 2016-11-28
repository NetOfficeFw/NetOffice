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
	/// Interface IMarkupContainer2 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IMarkupContainer2 : IMarkupContainer
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
                    _type = typeof(IMarkupContainer2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="ppChangeLog">NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog</param>
		/// <param name="fForward">Int32 fForward</param>
		/// <param name="fBackward">Int32 fBackward</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CreateChangeLog(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out NetOffice.MSHTMLApi.IHTMLChangeLog ppChangeLog, Int32 fForward, Int32 fBackward)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,false,false);
			ppChangeLog = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, ppChangeLog, fForward, fBackward);
			object returnItem = Invoker.MethodReturn(this, "CreateChangeLog", paramsArray);
			ppChangeLog = (NetOffice.MSHTMLApi.IHTMLChangeLog)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pChangeSink">NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink</param>
		/// <param name="pdwCookie">Int32 pdwCookie</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RegisterForDirtyRange(NetOffice.MSHTMLApi.IHTMLChangeSink pChangeSink, out Int32 pdwCookie)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pdwCookie = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pChangeSink, pdwCookie);
			object returnItem = Invoker.MethodReturn(this, "RegisterForDirtyRange", paramsArray);
			pdwCookie = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 UnRegisterForDirtyRange(Int32 dwCookie)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCookie);
			object returnItem = Invoker.MethodReturn(this, "UnRegisterForDirtyRange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="dwCookie">Int32 dwCookie</param>
		/// <param name="pIPointerBegin">NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin</param>
		/// <param name="pIPointerEnd">NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetAndClearDirtyRange(Int32 dwCookie, NetOffice.MSHTMLApi.IMarkupPointer pIPointerBegin, NetOffice.MSHTMLApi.IMarkupPointer pIPointerEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dwCookie, pIPointerBegin, pIPointerEnd);
			object returnItem = Invoker.MethodReturn(this, "GetAndClearDirtyRange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetVersionNumber()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetVersionNumber", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppElementMaster">NetOffice.MSHTMLApi.IHTMLElement ppElementMaster</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetMasterElement(out NetOffice.MSHTMLApi.IHTMLElement ppElementMaster)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppElementMaster = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppElementMaster);
			object returnItem = Invoker.MethodReturn(this, "GetMasterElement", paramsArray);
			ppElementMaster = (NetOffice.MSHTMLApi.IHTMLElement)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}