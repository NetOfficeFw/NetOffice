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
	/// DispatchInterface OMathMatCols 
	/// SupportByVersion Word, 12,14
	///</summary>
	[SupportByVersionAttribute("Word", 12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class OMathMatCols : COMObject ,IEnumerable<NetOffice.WordApi.OMathMatCol>
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
                    _type = typeof(OMathMatCols);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMatCols(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMatCols(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMatCols(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMatCols() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public OMathMatCols(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
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
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="index">Int32 Index</param>
		[SupportByVersionAttribute("Word", 12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.WordApi.OMathMatCol this[Int32 index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.WordApi.OMathMatCol newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMatCol;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		/// <param name="beforeCol">optional object BeforeCol</param>
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.OMathMatCol Add(object beforeCol)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(beforeCol);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.OMathMatCol newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMatCol;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Word 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Word", 12,14)]
		public NetOffice.WordApi.OMathMatCol Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.WordApi.OMathMatCol newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.WordApi.OMathMatCol.LateBindingApiWrapperType) as NetOffice.WordApi.OMathMatCol;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.WordApi.OMathMatCol> Member
        
        /// <summary>
		/// SupportByVersionAttribute Word, 12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
       public IEnumerator<NetOffice.WordApi.OMathMatCol> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.WordApi.OMathMatCol item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Word, 12,14
		/// </summary>
		[SupportByVersionAttribute("Word", 12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}