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
	/// Interface IHTMLEditServices 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IHTMLEditServices : COMObject
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
                    _type = typeof(IHTMLEditServices);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLEditServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices(string progId) : base(progId)
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
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AddDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIDesigner);
			object returnItem = Invoker.MethodReturn(this, "AddDesigner", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIDesigner">NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RemoveDesigner(NetOffice.MSHTMLApi.IHTMLEditDesigner pIDesigner)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIDesigner);
			object returnItem = Invoker.MethodReturn(this, "RemoveDesigner", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIContainer">NetOffice.MSHTMLApi.IMarkupContainer pIContainer</param>
		/// <param name="ppSelSvc">NetOffice.MSHTMLApi.ISelectionServices ppSelSvc</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetSelectionServices(NetOffice.MSHTMLApi.IMarkupContainer pIContainer, out NetOffice.MSHTMLApi.ISelectionServices ppSelSvc)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppSelSvc = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIContainer, ppSelSvc);
			object returnItem = Invoker.MethodReturn(this, "GetSelectionServices", paramsArray);
			ppSelSvc = (NetOffice.MSHTMLApi.ISelectionServices)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToSelectionAnchor(NetOffice.MSHTMLApi.IMarkupPointer pIStartAnchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIStartAnchor);
			object returnItem = Invoker.MethodReturn(this, "MoveToSelectionAnchor", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToSelectionEnd(NetOffice.MSHTMLApi.IMarkupPointer pIEndAnchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIEndAnchor);
			object returnItem = Invoker.MethodReturn(this, "MoveToSelectionEnd", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pStart">NetOffice.MSHTMLApi.IMarkupPointer pStart</param>
		/// <param name="pEnd">NetOffice.MSHTMLApi.IMarkupPointer pEnd</param>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SelectRange(NetOffice.MSHTMLApi.IMarkupPointer pStart, NetOffice.MSHTMLApi.IMarkupPointer pEnd, NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pStart, pEnd, eType);
			object returnItem = Invoker.MethodReturn(this, "SelectRange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}