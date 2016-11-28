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
	/// Interface ISelectionServicesListener 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class ISelectionServicesListener : COMObject
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
                    _type = typeof(ISelectionServicesListener);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISelectionServicesListener(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISelectionServicesListener(string progId) : base(progId)
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
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 BeginSelectionUndo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "BeginSelectionUndo", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 EndSelectionUndo()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "EndSelectionUndo", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pIElementStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementStart</param>
		/// <param name="pIElementEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd</param>
		/// <param name="pIElementContentStart">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart</param>
		/// <param name="pIElementContentEnd">NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 OnSelectedElementExit(NetOffice.MSHTMLApi.IMarkupPointer pIElementStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementEnd, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentStart, NetOffice.MSHTMLApi.IMarkupPointer pIElementContentEnd)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pIElementStart, pIElementEnd, pIElementContentStart, pIElementContentEnd);
			object returnItem = Invoker.MethodReturn(this, "OnSelectedElementExit", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="eType">NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType</param>
		/// <param name="pIListener">NetOffice.MSHTMLApi.ISelectionServicesListener pIListener</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 OnChangeType(NetOffice.MSHTMLApi.Enums._SELECTION_TYPE eType, NetOffice.MSHTMLApi.ISelectionServicesListener pIListener)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eType, pIListener);
			object returnItem = Invoker.MethodReturn(this, "OnChangeType", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="pTypeDetail">string pTypeDetail</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public Int32 GetTypeDetail(out string pTypeDetail)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true);
			pTypeDetail = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(pTypeDetail);
			object returnItem = Invoker.MethodReturn(this, "GetTypeDetail", paramsArray);
			pTypeDetail = (string)paramsArray[0];
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}