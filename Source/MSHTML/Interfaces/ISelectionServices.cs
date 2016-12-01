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
	/// Interface ISelectionServices 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class ISelectionServices : COMObject
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
                    _type = typeof(ISelectionServices);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISelectionServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServices(string progId) : base(progId)
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
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 SetSelectionType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eType, pIListener);
			object returnItem = Invoker.MethodReturn(this, "SetSelectionType", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppIContainer">NetOffice.MSHTMLApi.IMarkupContainer ppIContainer</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetMarkupContainer(out NetOffice.MSHTMLApi.IMarkupContainer ppIContainer)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppIContainer = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppIContainer);
			object returnItem = Invoker.MethodReturn(this, "GetMarkupContainer", paramsArray);
			ppIContainer = (NetOffice.MSHTMLApi.IMarkupContainer)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIStart">NetOffice.MSHTMLApi.IMarkupPointer pIStart</param>
		/// <param name="pIEnd">NetOffice.MSHTMLApi.IMarkupPointer pIEnd</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.ISegment ppISegmentAdded</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AddSegment(NetOffice.MSHTMLApi.IMarkupPointer pIStart, NetOffice.MSHTMLApi.IMarkupPointer pIEnd, out NetOffice.MSHTMLApi.ISegment ppISegmentAdded)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppISegmentAdded = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIStart, pIEnd, ppISegmentAdded);
			object returnItem = Invoker.MethodReturn(this, "AddSegment", paramsArray);
			ppISegmentAdded = (NetOffice.MSHTMLApi.ISegment)paramsArray[2];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIElement">NetOffice.MSHTMLApi.IHTMLElement pIElement</param>
		/// <param name="ppISegmentAdded">NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AddElementSegment(NetOffice.MSHTMLApi.IHTMLElement pIElement, out NetOffice.MSHTMLApi.IElementSegment ppISegmentAdded)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			ppISegmentAdded = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pIElement, ppISegmentAdded);
			object returnItem = Invoker.MethodReturn(this, "AddElementSegment", paramsArray);
			ppISegmentAdded = (NetOffice.MSHTMLApi.IElementSegment)paramsArray[1];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.ISegment pISegment</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RemoveSegment(NetOffice.MSHTMLApi.ISegment pISegment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pISegment);
			object returnItem = Invoker.MethodReturn(this, "RemoveSegment", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="ppISelectionServicesListener">NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetSelectionServicesListener(out NetOffice.MSHTMLApi.ISelectionServicesListener ppISelectionServicesListener)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			ppISelectionServicesListener = null;
			object[] paramsArray = Invoker.ValidateParamsArray(ppISelectionServicesListener);
			object returnItem = Invoker.MethodReturn(this, "GetSelectionServicesListener", paramsArray);
			ppISelectionServicesListener = (NetOffice.MSHTMLApi.ISelectionServicesListener)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}