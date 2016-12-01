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
	/// Interface IHTMLEditServices2 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IHTMLEditServices2 : IHTMLEditServices
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
                    _type = typeof(IHTMLEditServices2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLEditServices2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLEditServices2(string progId) : base(progId)
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
		/// <param name="pIStartAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToSelectionAnchorEx(NetOffice.MSHTMLApi.IDisplayPointer pIStartAnchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIStartAnchor);
			object returnItem = Invoker.MethodReturn(this, "MoveToSelectionAnchorEx", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIEndAnchor">NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 MoveToSelectionEndEx(NetOffice.MSHTMLApi.IDisplayPointer pIEndAnchor)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIEndAnchor);
			object returnItem = Invoker.MethodReturn(this, "MoveToSelectionEndEx", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fReCompute">Int32 fReCompute</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 FreezeVirtualCaretPos(Int32 fReCompute)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fReCompute);
			object returnItem = Invoker.MethodReturn(this, "FreezeVirtualCaretPos", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="fReset">Int32 fReset</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 UnFreezeVirtualCaretPos(Int32 fReset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fReset);
			object returnItem = Invoker.MethodReturn(this, "UnFreezeVirtualCaretPos", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}