using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSHTMLApi
{
	///<summary>
	/// DispatchInterface IHTMLDOMTextNode2 
	/// SupportByVersion MSHTML, 4
	///</summary>
	[SupportByVersionAttribute("MSHTML", 4)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IHTMLDOMTextNode2 : IHTMLDOMTextNode
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
                    _type = typeof(IHTMLDOMTextNode2);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IHTMLDOMTextNode2(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IHTMLDOMTextNode2(string progId) : base(progId)
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
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 Count</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public string substringData(Int32 offset, Int32 count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(offset, count);
			object returnItem = Invoker.MethodReturn(this, "substringData", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void appendData(string bstrstring)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrstring);
			Invoker.Method(this, "appendData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void insertData(Int32 offset, string bstrstring)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(offset, bstrstring);
			Invoker.Method(this, "insertData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 Count</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void deleteData(Int32 offset, Int32 count)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(offset, count);
			Invoker.Method(this, "deleteData", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSHTML 4
		/// 
		/// </summary>
		/// <param name="offset">Int32 offset</param>
		/// <param name="count">Int32 Count</param>
		/// <param name="bstrstring">string bstrstring</param>
		[SupportByVersionAttribute("MSHTML", 4)]
		public void replaceData(Int32 offset, Int32 count, string bstrstring)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(offset, count, bstrstring);
			Invoker.Method(this, "replaceData", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}