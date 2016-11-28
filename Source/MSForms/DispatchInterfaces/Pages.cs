using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSFormsApi
{
	///<summary>
	/// DispatchInterface Pages 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Pages : COMObject ,IEnumerable<object>
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
                    _type = typeof(Pages);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Pages(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Pages(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
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
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersionAttribute("MSForms", 2)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public object this[object varg]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(varg);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public object Enum()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Enum", paramsArray);
			object newObject = Factory.CreateObjectFromComProxy(this,returnItem);
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		/// <param name="lIndex">optional object lIndex</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption, object lIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName, bstrCaption, lIndex);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrName">optional object bstrName</param>
		/// <param name="bstrCaption">optional object bstrCaption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page Add(object bstrName, object bstrCaption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrName, bstrCaption);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page _AddCtrl(Int32 clsid, string bstrName, string bstrCaption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(clsid, bstrName, bstrCaption);
			object returnItem = Invoker.MethodReturn(this, "_AddCtrl", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		/// <param name="bstrName">string bstrName</param>
		/// <param name="bstrCaption">string bstrCaption</param>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Page _InsertCtrl(Int32 clsid, string bstrName, string bstrCaption, Int32 lIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(clsid, bstrName, bstrCaption, lIndex);
			object returnItem = Invoker.MethodReturn(this, "_InsertCtrl", paramsArray);
			NetOffice.MSFormsApi.Page newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Page.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Page;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByIndex(Int32 lIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lIndex);
			object returnItem = Invoker.MethodReturn(this, "_GetItemByIndex", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="pstrName">string pstrName</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByName(string pstrName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pstrName);
			object returnItem = Invoker.MethodReturn(this, "_GetItemByName", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="varg">object varg</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Remove(object varg)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(varg);
			Invoker.Method(this, "Remove", paramsArray);
		}

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

		#endregion

       #region IEnumerable<object> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSForms, 2
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
       public IEnumerator<object> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (object item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSForms, 2
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}