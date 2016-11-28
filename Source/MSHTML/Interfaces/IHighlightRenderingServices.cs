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
	/// Interface IHighlightRenderingServices 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IHighlightRenderingServices : COMObject
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
                    _type = typeof(IHighlightRenderingServices);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHighlightRenderingServices(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHighlightRenderingServices(string progId) : base(progId)
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
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		/// <param name="pIRenderStyle">NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle</param>
		/// <param name="ppISegment">NetOffice.MSHTMLApi.IHighlightSegment ppISegment</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 AddSegment(NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd, NetOffice.MSHTMLApi.IHTMLRenderStyle pIRenderStyle, out NetOffice.MSHTMLApi.IHighlightSegment ppISegment)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,true);
			ppISegment = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pDispPointerStart, pDispPointerEnd, pIRenderStyle, ppISegment);
			object returnItem = Invoker.MethodReturn(this, "AddSegment", paramsArray);
			ppISegment = (NetOffice.MSHTMLApi.IHighlightSegment)paramsArray[3];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		/// <param name="pDispPointerStart">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart</param>
		/// <param name="pDispPointerEnd">NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveSegmentToPointers(NetOffice.MSHTMLApi.IHighlightSegment pISegment, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerStart, NetOffice.MSHTMLApi.IDisplayPointer pDispPointerEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pISegment, pDispPointerStart, pDispPointerEnd);
			object returnItem = Invoker.MethodReturn(this, "MoveSegmentToPointers", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pISegment">NetOffice.MSHTMLApi.IHighlightSegment pISegment</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 RemoveSegment(NetOffice.MSHTMLApi.IHighlightSegment pISegment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pISegment);
			object returnItem = Invoker.MethodReturn(this, "RemoveSegment", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}