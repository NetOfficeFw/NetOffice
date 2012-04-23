using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.WordApi
{
	///<summary>
	/// DispatchInterface StyleSheets 
	/// SupportByVersion Word, 10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class StyleSheets : COMObject ,IEnumerable<NetOffice.WordApi.StyleSheet>
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
                    _type = typeof(StyleSheets);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StyleSheets(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StyleSheets(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StyleSheets(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StyleSheets() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public StyleSheets(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
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
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
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
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
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
		/// SupportByVersion Word 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.StyleSheet this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.StyleSheet newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.StyleSheet.LateBindingApiWrapperType) as NetOffice.WordApi.StyleSheet;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 10, 11, 12, 14
		/// </summary>
		/// <param name="fileName">string FileName</param>
		/// <param name="linkType">NetOffice.WordApi.Enums.WdStyleSheetLinkType LinkType</param>
		/// <param name="title">string Title</param>
		/// <param name="precedence">NetOffice.WordApi.Enums.WdStyleSheetPrecedence Precedence</param>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		public NetOffice.WordApi.StyleSheet Add(string fileName, NetOffice.WordApi.Enums.WdStyleSheetLinkType linkType, string title, NetOffice.WordApi.Enums.WdStyleSheetPrecedence precedence)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(fileName, linkType, title, precedence);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.StyleSheet newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.StyleSheet.LateBindingApiWrapperType) as NetOffice.WordApi.StyleSheet;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.StyleSheet> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
       public IEnumerator<NetOffice.WordApi.StyleSheet> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.StyleSheet item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 10,11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 10,11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}