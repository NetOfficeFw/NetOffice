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
	/// Interface IDisplayServices 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IDisplayServices : COMObject
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
                    _type = typeof(IDisplayServices);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDisplayServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDisplayServices(string progId) : base(progId)
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
		/// <param name="ppDispPointer">NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 CreateDisplayPointer(out NetOffice.MSHTMLApi.IDisplayPointer ppDispPointer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppDispPointer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppDispPointer);
			object returnItem = Invoker.MethodReturn(this, "CreateDisplayPointer", paramsArray);
			ppDispPointer = (NetOffice.MSHTMLApi.IDisplayPointer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pRect">tagRECT pRect</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 TransformRect(tagRECT pRect, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pRect, eSource, eDestination, pIElement);
			object returnItem = Invoker.MethodReturn(this, "TransformRect", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="eSource">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource</param>
		/// <param name="eDestination">NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination</param>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 TransformPoint(tagPOINT pPoint, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eSource, NetOffice.MSHTMLApi.Enums._COORD_SYSTEM eDestination, NetOffice.MSHTMLApi.IHTMLElement pIElement)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pPoint, eSource, eDestination, pIElement);
			object returnItem = Invoker.MethodReturn(this, "TransformPoint", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppCaret">NetOffice.MSHTMLApi.IHTMLCaret ppCaret</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCaret(out NetOffice.MSHTMLApi.IHTMLCaret ppCaret)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppCaret = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppCaret);
			object returnItem = Invoker.MethodReturn(this, "GetCaret", paramsArray);
			ppCaret = (NetOffice.MSHTMLApi.IHTMLCaret)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPointer">NetOffice.MSHTMLApi.IMarkupPointer pPointer</param>
		/// <param name="ppComputedStyle">NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetComputedStyle(NetOffice.MSHTMLApi.IMarkupPointer pPointer, out NetOffice.MSHTMLApi.IHTMLComputedStyle ppComputedStyle)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppComputedStyle = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pPointer, ppComputedStyle);
			object returnItem = Invoker.MethodReturn(this, "GetComputedStyle", paramsArray);
			ppComputedStyle = (NetOffice.MSHTMLApi.IHTMLComputedStyle)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="rect">tagRECT rect</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 ScrollRectIntoView(NetOffice.MSHTMLApi.IHTMLElement pIElement, tagRECT rect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIElement, rect);
			object returnItem = Invoker.MethodReturn(this, "ScrollRectIntoView", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="pfHasFlowLayout">Int32 pfHasFlowLayout</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 HasFlowLayout(NetOffice.MSHTMLApi.IHTMLElement pIElement, out Int32 pfHasFlowLayout)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			pfHasFlowLayout = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pIElement, pfHasFlowLayout);
			object returnItem = Invoker.MethodReturn(this, "HasFlowLayout", paramsArray);
			pfHasFlowLayout = (Int32)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}