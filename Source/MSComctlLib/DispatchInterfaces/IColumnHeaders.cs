using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSComctlLibApi
{
	///<summary>
	/// DispatchInterface IColumnHeaders 
	/// SupportByVersion MSComctlLib, 6
	///</summary>
	[SupportByVersionAttribute("MSComctlLib", 6)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IColumnHeaders : COMObject ,IEnumerable<NetOffice.MSComctlLibApi.IColumnHeader>
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
                    _type = typeof(IColumnHeaders);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IColumnHeaders(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IColumnHeaders(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Count", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.IColumnHeader get_ControlDefault(object index)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "ControlDefault", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader ControlDefault(object index)
		{
			return get_ControlDefault(index);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSComctlLibApi.IColumnHeader this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		/// <param name="width">optional object Width</param>
		/// <param name="alignment">optional object Alignment</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width, object alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text, width, alignment);
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key);
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text);
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		/// <param name="width">optional object Width</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add_PreVB98(object index, object key, object text, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text, width);
			object returnItem = Invoker.MethodReturn(this, "Add_PreVB98", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public void Clear()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Clear", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public void Remove(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			Invoker.Method(this, "Remove", paramsArray);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		/// <param name="width">optional object Width</param>
		/// <param name="alignment">optional object Alignment</param>
		/// <param name="icon">optional object Icon</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment, object icon)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text, width, alignment, icon);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		/// <param name="width">optional object Width</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text, width);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="index">optional object Index</param>
		/// <param name="key">optional object Key</param>
		/// <param name="text">optional object Text</param>
		/// <param name="width">optional object Width</param>
		/// <param name="alignment">optional object Alignment</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.IColumnHeader Add(object index, object key, object text, object width, object alignment)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(index, key, text, width, alignment);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.IColumnHeader newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.IColumnHeader;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSComctlLibApi.IColumnHeader> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSComctlLib, 6
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
       public IEnumerator<NetOffice.MSComctlLibApi.IColumnHeader> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSComctlLibApi.IColumnHeader item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSComctlLib, 6
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}