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
	/// DispatchInterface XSLTransforms 
	/// SupportByVersion Word, 11,12,14
	///</summary>
	[SupportByVersionAttribute("Word", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class XSLTransforms : COMObject ,IEnumerable<NetOffice.WordApi.XSLTransform>
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
                    _type = typeof(XSLTransforms);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XSLTransforms(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XSLTransforms(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XSLTransforms(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XSLTransforms() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public XSLTransforms(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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
		/// SupportByVersion Word 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
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

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Word", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.XSLTransform this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.XSLTransform newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XSLTransform.LateBindingApiWrapperType) as NetOffice.WordApi.XSLTransform;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="location">string Location</param>
		/// <param name="alias">object Alias</param>
		/// <param name="installForAllUsers">optional bool InstallForAllUsers = false</param>
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.XSLTransform Add(string location, object alias, bool installForAllUsers)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(location, alias, installForAllUsers);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.XSLTransform newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XSLTransform.LateBindingApiWrapperType) as NetOffice.WordApi.XSLTransform;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 11, 12, 14
		/// </summary>
		/// <param name="location">string Location</param>
		/// <param name="alias">object Alias</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 11,12,14)]
		public NetOffice.WordApi.XSLTransform Add(string location, object alias)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(location, alias);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.XSLTransform newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.XSLTransform.LateBindingApiWrapperType) as NetOffice.WordApi.XSLTransform;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.XSLTransform> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
       public IEnumerator<NetOffice.WordApi.XSLTransform> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.XSLTransform item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}