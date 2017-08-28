using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSProjectApi
{
	/// <summary>
	/// DispatchInterface ViewsSingle 
	/// SupportByVersion MSProject, 11,12,14
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff920756(v=office.14).aspx </remarks>
	[SupportByVersion("MSProject", 11,12,14)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class ViewsSingle : Views
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

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
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ViewsSingle(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

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
		
		/// <param name="progId">registered progID</param>
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
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		/// <param name="highlightFilt">optional bool HighlightFilt = false</param>
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group, object highlightFilt)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, new object[]{ name, screen, showInMenu, table, filter, group, highlightFilt });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, name);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, name, screen);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, name, screen, showInMenu);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, name, screen, showInMenu, table);
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, new object[]{ name, screen, showInMenu, table, filter });
		}

		/// <summary>
		/// SupportByVersion MSProject 11, 12, 14
		/// </summary>
		/// <param name="name">string name</param>
		/// <param name="screen">optional NetOffice.MSProjectApi.Enums.PjViewScreen Screen = 1</param>
		/// <param name="showInMenu">optional bool ShowInMenu = false</param>
		/// <param name="table">optional object table</param>
		/// <param name="filter">optional object filter</param>
		/// <param name="group">optional object group</param>
		[CustomMethod]
		[SupportByVersion("MSProject", 11,12,14)]
		public NetOffice.MSProjectApi.ViewSingle Add(string name, object screen, object showInMenu, object table, object filter, object group)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.MSProjectApi.ViewSingle>(this, "Add", NetOffice.MSProjectApi.ViewSingle.LateBindingApiWrapperType, new object[]{ name, screen, showInMenu, table, filter, group });
		}

		#endregion

		#pragma warning restore
	}
}
