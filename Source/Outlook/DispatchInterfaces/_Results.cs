using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
	/// DispatchInterface _Results 
	/// SupportByVersion Outlook, 10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Outlook", 10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Results : COMObject ,IEnumerable<object>
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
                    _type = typeof(_Results);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Results(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Results(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Results(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Results() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Results(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public NetOffice.OutlookApi._Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.OutlookApi._Application newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public NetOffice.OutlookApi.Enums.OlObjectClass Class
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Class", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlObjectClass)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public NetOffice.OutlookApi._NameSpace Session
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Session", paramsArray);
				NetOffice.OutlookApi._NameSpace newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.OutlookApi._NameSpace;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
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
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
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
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object RawTable
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RawTable", paramsArray);
				COMObject newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public NetOffice.OutlookApi.Enums.OlItemType DefaultItemType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultItemType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OutlookApi.Enums.OlItemType)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultItemType", paramsArray);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				object newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public object GetFirst()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetFirst", paramsArray);
			object newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public object GetLast()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetLast", paramsArray);
			object newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public object GetNext()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetNext", paramsArray);
			object newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public object GetPrevious()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "GetPrevious", paramsArray);
			object newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public void ResetColumns()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ResetColumns", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		/// <param name="columns">string Columns</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public void SetColumns(string columns)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(columns);
			Invoker.Method(this, "SetColumns", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		/// <param name="property">string Property</param>
		/// <param name="descending">optional object Descending</param>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public void Sort(string property, object descending)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(property, descending);
			Invoker.Method(this, "Sort", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Outlook 10, 11, 12, 14
		/// </summary>
		/// <param name="property">string Property</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
		public void Sort(string property)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(property);
			Invoker.Method(this, "Sort", paramsArray);
		}

		#endregion
       #region IEnumerable<object> Member
        
        /// <summary>
		/// SupportByVersionAttribute Outlook, 10,11,12,14
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
        [CustomEnumerator]
       public IEnumerator<object> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (object item in innerEnumerator)
               yield return item;
       }

       #endregion
   
       #region IEnumerable Members
        
       /// <summary>
		/// SupportByVersionAttribute Outlook, 10,11,12,14
		/// This is a custom enumerator from NetOffice
		/// </summary>
		[SupportByVersionAttribute("Outlook", 10,11,12,14)]
        [CustomEnumerator]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
       {
            int count = Count;
            object[] enumeratorObjects = new object[count];
            for (int i = 0; i < count; i++)
                enumeratorObjects[i] = this[i+1];

            foreach (object item in enumeratorObjects)
                yield return item;
       }

       #endregion
       		#pragma warning restore
	}
}