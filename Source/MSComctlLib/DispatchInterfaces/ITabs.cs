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
	/// DispatchInterface ITabs 
	/// SupportByVersion MSComctlLib, 6
	///</summary>
	[SupportByVersionAttribute("MSComctlLib", 6)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ITabs : COMObject ,IEnumerable<NetOffice.MSComctlLibApi.ITab>
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
                    _type = typeof(ITabs);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ITabs(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ITabs(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public Int16 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt16(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Count", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSComctlLibApi.ITab get_ControlDefault(object pvIndex)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			object returnItem = Invoker.PropertyGet(this, "ControlDefault", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_ControlDefault(object pvIndex, NetOffice.MSComctlLibApi.ITab value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			Invoker.PropertySet(this, "ControlDefault", paramsArray, value);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Alias for get_ControlDefault
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab ControlDefault(object pvIndex)
		{
			return get_ControlDefault(pvIndex);
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// Get/Set
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSComctlLibApi.ITab this[object pvIndex]
		{
			get
			{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
			}
			set
			{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			Invoker.PropertySet(this, "Item", paramsArray, value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="pvIndex">object pvIndex</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public void Remove(object pvIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			Invoker.Method(this, "Remove", paramsArray);
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
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		/// <param name="pvImage">optional object pvImage</param>
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption, object pvImage)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex, pvKey, pvCaption, pvImage);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex, pvKey);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSComctlLib 6
		/// 
		/// </summary>
		/// <param name="pvIndex">optional object pvIndex</param>
		/// <param name="pvKey">optional object pvKey</param>
		/// <param name="pvCaption">optional object pvCaption</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSComctlLib", 6)]
		public NetOffice.MSComctlLibApi.ITab Add(object pvIndex, object pvKey, object pvCaption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pvIndex, pvKey, pvCaption);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSComctlLibApi.ITab newObject = Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.MSComctlLibApi.ITab;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSComctlLibApi.ITab> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSComctlLib, 6
		/// </summary>
		[SupportByVersionAttribute("MSComctlLib", 6)]
       public IEnumerator<NetOffice.MSComctlLibApi.ITab> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSComctlLibApi.ITab item in innerEnumerator)
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