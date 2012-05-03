using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _DependencyInfo 
	/// SupportByVersion Access, 11,12,14
	///</summary>
	[SupportByVersionAttribute("Access", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _DependencyInfo : COMObject
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
                    _type = typeof(_DependencyInfo);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DependencyInfo(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DependencyInfo(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DependencyInfo(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DependencyInfo() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _DependencyInfo(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
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
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public NetOffice.AccessApi._DependencyObjects Dependants
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dependants", paramsArray);
				NetOffice.AccessApi._DependencyObjects newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._DependencyObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public NetOffice.AccessApi._DependencyObjects Dependencies
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Dependencies", paramsArray);
				NetOffice.AccessApi._DependencyObjects newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._DependencyObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public NetOffice.AccessApi._DependencyObjects OutOfDateObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OutOfDateObjects", paramsArray);
				NetOffice.AccessApi._DependencyObjects newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._DependencyObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public NetOffice.AccessApi._DependencyObjects InsufficientPermissions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "InsufficientPermissions", paramsArray);
				NetOffice.AccessApi._DependencyObjects newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._DependencyObjects;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public NetOffice.AccessApi._DependencyObjects UnsupportedObjects
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UnsupportedObjects", paramsArray);
				NetOffice.AccessApi._DependencyObjects newObject = NetOffice.Factory.CreateObjectFromComProxy(this,returnItem) as NetOffice.AccessApi._DependencyObjects;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("Access", 11,12,14)]
		public bool IsMemberSafe(Int32 dispid)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispid);
			object returnItem = Invoker.MethodReturn(this, "IsMemberSafe", paramsArray);
			return (bool)returnItem;
		}

		#endregion
		#pragma warning restore
	}
}