using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.DAOApi
{
	///<summary>
	/// DispatchInterface _Group 
	/// SupportByVersion DAO, 5,12
	///</summary>
	[SupportByVersionAttribute("DAO", 5,12)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _Group : _DAO
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
                    _type = typeof(_Group);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Group(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Group(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Group(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Group() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _Group(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 5,12)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "Name", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("DAO", 5,12)]
		public string PID
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PID", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "PID", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// Get
		/// </summary>
		[SupportByVersionAttribute("DAO", 5,12)]
		public NetOffice.DAOApi.Users Users
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Users", paramsArray);
				NetOffice.DAOApi.Users newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.DAOApi.Users.LateBindingApiWrapperType) as NetOffice.DAOApi.Users;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="pID">optional object PID</param>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("DAO", 5,12)]
		public NetOffice.DAOApi.User CreateUser(object name, object pID, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, pID, password);
			object returnItem = Invoker.MethodReturn(this, "CreateUser", paramsArray);
			NetOffice.DAOApi.User newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.User.LateBindingApiWrapperType) as NetOffice.DAOApi.User;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 5,12)]
		public NetOffice.DAOApi.User CreateUser()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "CreateUser", paramsArray);
			NetOffice.DAOApi.User newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.User.LateBindingApiWrapperType) as NetOffice.DAOApi.User;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// </summary>
		/// <param name="name">optional object Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 5,12)]
		public NetOffice.DAOApi.User CreateUser(object name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "CreateUser", paramsArray);
			NetOffice.DAOApi.User newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.User.LateBindingApiWrapperType) as NetOffice.DAOApi.User;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion DAO 5, 12
		/// </summary>
		/// <param name="name">optional object Name</param>
		/// <param name="pID">optional object PID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("DAO", 5,12)]
		public NetOffice.DAOApi.User CreateUser(object name, object pID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, pID);
			object returnItem = Invoker.MethodReturn(this, "CreateUser", paramsArray);
			NetOffice.DAOApi.User newObject = NetOffice.Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.DAOApi.User.LateBindingApiWrapperType) as NetOffice.DAOApi.User;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}