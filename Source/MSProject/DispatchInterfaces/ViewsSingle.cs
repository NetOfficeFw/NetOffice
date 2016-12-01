using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.MSProjectApi
{
	///<summary>
	/// DispatchInterface ViewsSingle 
	/// SupportByVersion MSProject, 11,12,14
	/// MSDN Online Documentation: http://msdn.microsoft.com/en-us/en-us/library/office/ff920756(v=office.14).aspx
	///</summary>
	[SupportByVersionAttribute("MSProject", 11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class ViewsSingle : Views
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
                    _type = typeof(ViewsSingle);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ViewsSingle(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ViewsSingle(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object Table</param>
		/// <param name="filter">optional object Filter</param>
		/// <param name="group">optional object Group</param>
		/// <param name="highlightFilt">optional bool HighlightFilt = false</param>
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group, object highlightFilt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen, showInMenu, table, filter, group, highlightFilt);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen, showInMenu);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object Table</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen, showInMenu, table);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object Table</param>
		/// <param name="filter">optional object Filter</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen, showInMenu, table, filter);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// 
		/// </summary>
		/// <param name="name">string Name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object Table</param>
		/// <param name="filter">optional object Filter</param>
		/// <param name="group">optional object Group</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(name, screen, showInMenu, table, filter, group);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.MSProjectApi.ViewSingle newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType) as NetOffice.MSProjectApi.ViewSingle;
			return newObject;
		}

		#endregion
		#pragma warning restore
	}
}