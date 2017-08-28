using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface HTMLProject 
	/// SupportByVersion Office, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Office", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class HTMLProject : _IMsoDispObj
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
                    _type = typeof(HTMLProject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public HTMLProject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public HTMLProject(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public HTMLProject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoHTMLProjectState State
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoHTMLProjectState>(this, "State");
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public NetOffice.OfficeApi.HTMLProjectItems HTMLProjectItems
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.HTMLProjectItems>(this, "HTMLProjectItems", NetOffice.OfficeApi.HTMLProjectItems.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Office", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="refresh">optional bool Refresh = true</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void RefreshProject(object refresh)
		{
			 Factory.ExecuteMethod(this, "RefreshProject", refresh);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void RefreshProject()
		{
			 Factory.ExecuteMethod(this, "RefreshProject");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="refresh">optional bool Refresh = true</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void RefreshDocument(object refresh)
		{
			 Factory.ExecuteMethod(this, "RefreshDocument", refresh);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void RefreshDocument()
		{
			 Factory.ExecuteMethod(this, "RefreshDocument");
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="openKind">optional NetOffice.OfficeApi.Enums.MsoHTMLProjectOpen OpenKind = 0</param>
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Open(object openKind)
		{
			 Factory.ExecuteMethod(this, "Open", openKind);
		}

		/// <summary>
		/// SupportByVersion Office 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		[CustomMethod]
		[SupportByVersion("Office", 9,10,11,12,14,15,16)]
		public void Open()
		{
			 Factory.ExecuteMethod(this, "Open");
		}

		#endregion

		#pragma warning restore
	}
}
