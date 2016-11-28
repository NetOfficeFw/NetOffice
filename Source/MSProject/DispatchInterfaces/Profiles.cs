using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface Profiles 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920661(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class Profiles : COMObject ,IEnumerable<NetOffice.MSProjectApi.Profile>
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
                    _type = typeof(Profiles);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Profiles(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Profiles(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.MSProjectApi.Profile this[object index]
		{
			get
{			
			object[] paramsArray = Invoker.ValidateParamsArray(index);
			object returnItem = Invoker.PropertyGet(this, "Item", paramsArray);
			NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
			return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
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
		/// SupportByVersion MSProject 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Profile DefaultProfile
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultProfile", paramsArray);
				NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
				return newObject;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultProfile", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Profile ActiveProfile
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ActiveProfile", paramsArray);
				NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="server">string Server</param>
		/// <param name="loginType">optional NetOffice.MSProjectApi.Enums.PjLoginType LoginType = 1</param>
		/// <param name="userName">optional string UserName = </param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Profile Add(string name, string server, object loginType, object userName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, server, loginType, userName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="server">string Server</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Profile Add(string name, string server)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, server);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="server">string Server</param>
		/// <param name="loginType">optional NetOffice.MSProjectApi.Enums.PjLoginType LoginType = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.Profile Add(string name, string server, object loginType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, server, loginType);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.Profile newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.Profile.LateBindingApiWrapperType) as NetOffice.MSProjectApi.Profile;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.MSProjectApi.Profile> Member
        
        /// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
       public IEnumerator<NetOffice.MSProjectApi.Profile> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.MSProjectApi.Profile item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute MSProject, 11,12,14
		/// </summary>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion
		#pragma warning restore
	}
}