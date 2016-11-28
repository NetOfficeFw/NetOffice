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
	/// Interface IHTMLCaret 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IHTMLCaret : COMObject
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
                    _type = typeof(IHTMLCaret);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLCaret(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLCaret(string progId) : base(progId)
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
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveCaretToPointer(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer, fScrollIntoView, eDir);
			object returnItem = Invoker.MethodReturn(this, "MoveCaretToPointer", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		/// <param name="fVisible">Int32 fVisible</param>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveCaretToPointerEx(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer, Int32 fVisible, Int32 fScrollIntoView, NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer, fVisible, fScrollIntoView, eDir);
			object returnItem = Invoker.MethodReturn(this, "MoveCaretToPointerEx", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIMarkupPointer">NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveMarkupPointerToCaret(NetOffice.MSHTMLApi.IMarkupPointer pIMarkupPointer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIMarkupPointer);
			object returnItem = Invoker.MethodReturn(this, "MoveMarkupPointerToCaret", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pDispPointer">NetOffice.MSHTMLApi.IDisplayPointer pDispPointer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveDisplayPointerToCaret(NetOffice.MSHTMLApi.IDisplayPointer pDispPointer)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointer);
			object returnItem = Invoker.MethodReturn(this, "MoveDisplayPointerToCaret", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIsVisible">Int32 pIsVisible</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 IsVisible(out Int32 pIsVisible)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pIsVisible = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(pIsVisible);
			object returnItem = Invoker.MethodReturn(this, "IsVisible", paramsArray);
			pIsVisible = (Int32)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fScrollIntoView">Int32 fScrollIntoView</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Show(Int32 fScrollIntoView)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fScrollIntoView);
			object returnItem = Invoker.MethodReturn(this, "Show", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 Hide()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Hide", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pText">Int16 pText</param>
		/// <param name="lLen">Int32 lLen</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 InsertText(Int16 pText, Int32 lLen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pText, lLen);
			object returnItem = Invoker.MethodReturn(this, "InsertText", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 scrollIntoView()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "scrollIntoView", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pPoint">tagPOINT pPoint</param>
		/// <param name="fTranslate">Int32 fTranslate</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetLocation(out tagPOINT pPoint, Int32 fTranslate)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,false);
			pPoint = new NetOffice.MSHTMLApi.tagPOINT();
			object[] paramsArray = Invoker.ValidateParamsArray(pPoint, fTranslate);
			object returnItem = Invoker.MethodReturn(this, "GetLocation", paramsArray);
			pPoint = (tagPOINT)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="peDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetCaretDirection(out NetOffice.MSHTMLApi.Enums._CARET_DIRECTION peDir)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			peDir = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(peDir);
			object returnItem = Invoker.MethodReturn(this, "GetCaretDirection", paramsArray);
			peDir = (NetOffice.MSHTMLApi.Enums._CARET_DIRECTION)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="eDir">NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetCaretDirection(NetOffice.MSHTMLApi.Enums._CARET_DIRECTION eDir)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eDir);
			object returnItem = Invoker.MethodReturn(this, "SetCaretDirection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}