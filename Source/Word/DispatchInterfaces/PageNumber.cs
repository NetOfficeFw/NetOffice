using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface PageNumber 
	/// SupportByVersion Word, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class PageNumber : COMObject
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
                    _type = typeof(PageNumber);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageNumber(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageNumber(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageNumber(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageNumber() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PageNumber(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.WordApi.Application newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.WordApi.Application.LateBindingApiWrapperType) as NetOffice.WordApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public Int32 Index
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Index", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public NetOffice.WordApi.Enums.WdPageNumberAlignment Alignment
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Alignment", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.WordApi.Enums.WdPageNumberAlignment)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Alignment", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Select()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Select", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Word 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Word", 9,10,11,12,14)]
		public void Delete()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Delete", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}