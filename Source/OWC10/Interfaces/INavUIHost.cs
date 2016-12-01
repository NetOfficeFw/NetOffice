using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// Interface INavUIHost 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class INavUIHost : COMObject
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
                    _type = typeof(INavUIHost);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public INavUIHost(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public INavUIHost(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 IsButtonEnabled(Int32 navbtn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(navbtn);
			object returnItem = Invoker.MethodReturn(this, "IsButtonEnabled", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		/// <param name="cancel">Int32 Cancel</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 BeforeButtonClick(Int32 navbtn, Int32 cancel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(navbtn, cancel);
			object returnItem = Invoker.MethodReturn(this, "BeforeButtonClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="navbtn">Int32 navbtn</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 AfterButtonClick(Int32 navbtn)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(navbtn);
			object returnItem = Invoker.MethodReturn(this, "AfterButtonClick", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="displayText">string DisplayText</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetDisplayText(string displayText)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayText);
			object returnItem = Invoker.MethodReturn(this, "GetDisplayText", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 OnNavUIChange()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OnNavUIChange", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 IsFilterOn()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "IsFilterOn", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 IsContextBiDi()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "IsContextBiDi", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="fontName">string FontName</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetFontName(string fontName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fontName);
			object returnItem = Invoker.MethodReturn(this, "GetFontName", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}