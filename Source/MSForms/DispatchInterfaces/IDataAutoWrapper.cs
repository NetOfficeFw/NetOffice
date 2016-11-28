using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSFormsApi
{
	///<summary>
	/// DispatchInterface IDataAutoWrapper 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IDataAutoWrapper : COMObject
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
                    _type = typeof(IDataAutoWrapper);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDataAutoWrapper(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataAutoWrapper(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Clear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Clear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="format">object Format</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public bool GetFormat(object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			object returnItem = Invoker.MethodReturn(this, "GetFormat", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="format">optional object Format</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public string GetText(object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(format);
			object returnItem = Invoker.MethodReturn(this, "GetText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public string GetText()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetText", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		/// <param name="format">optional object Format</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void SetText(string text, object format)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text, format);
			Invoker.Method(this, "SetText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="text">string Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public void SetText(string text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(text);
			Invoker.Method(this, "SetText", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void PutInClipboard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "PutInClipboard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void GetFromClipboard()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "GetFromClipboard", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="oKEffect">optional object OKEffect</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag(object oKEffect)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oKEffect);
			object returnItem = Invoker.MethodReturn(this, "StartDrag", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.MSFormsApi.Enums.fmDropEffect)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDropEffect StartDrag()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "StartDrag", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.MSFormsApi.Enums.fmDropEffect)intReturnItem;
		}

		#endregion
		#pragma warning restore
	}
}