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
	/// DispatchInterface Controls 
	/// SupportByVersion MSForms, 2
	///</summary>
	[SupportByVersionAttribute("MSForms", 2)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Controls : COMObject ,IEnumerable<object>
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
                    _type = typeof(Controls);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Controls(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Controls(string progId) : base(progId)
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
		public void Clear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Clear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="cx">Int32 cx</param>
		/// <param name="cy">Int32 cy</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void _Move(Int32 cx, Int32 cy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx, cy);
			Invoker.Method(this, "_Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void SelectAll()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SelectAll", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="clsid">Int32 clsid</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _AddByClass(Int32 clsid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(clsid);
			object returnItem = Invoker.MethodReturn(this, "_AddByClass", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void AlignToGrid()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AlignToGrid", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void BringForward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringForward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void BringToFront()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "BringToFront", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Copy()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Copy", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSForms", 2)]
		public void Cut()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Cut", paramsArray);
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
		/// <param name="pstr">string pstr</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByName(string pstr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pstr);
			object returnItem = Invoker.MethodReturn(this, "_GetItemByName", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="iD">Int32 ID</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control _GetItemByID(Int32 iD)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(iD);
			object returnItem = Invoker.MethodReturn(this, "_GetItemByID", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void SendBackward()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendBackward", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSForms", 2)]
		public void SendToBack()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SendToBack", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="cx">Single cx</param>
		/// <param name="cy">Single cy</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public void Move(Single cx, Single cy)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(cx, cy);
			Invoker.Method(this, "Move", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object Name</param>
		/// <param name="visible">optional object Visible</param>
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID, object name, object visible)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrProgID, name, visible);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrProgID);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSFormsApi.Control newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSFormsApi.Control.LateBindingApiWrapperType) as NetOffice.MSFormsApi.Control;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// 
		/// </summary>
		/// <param name="bstrProgID">string bstrProgID</param>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSForms", 2)]
		public NetOffice.MSFormsApi.Control Add(string bstrProgID, object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrProgID, name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
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