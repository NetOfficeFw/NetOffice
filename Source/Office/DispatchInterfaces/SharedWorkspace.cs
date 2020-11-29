using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface SharedWorkspace 
	/// SupportByVersion Office, 11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace"/> </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class SharedWorkspace : _IMsoDispObj
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
                    _type = typeof(SharedWorkspace);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public SharedWorkspace(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public SharedWorkspace(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public SharedWorkspace(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Name"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Name", value);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Members"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceMembers Members
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceMembers>(this, "Members", NetOffice.OfficeApi.SharedWorkspaceMembers.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Tasks"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceTasks Tasks
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceTasks>(this, "Tasks", NetOffice.OfficeApi.SharedWorkspaceTasks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Files"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceFiles Files
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceFiles>(this, "Files", NetOffice.OfficeApi.SharedWorkspaceFiles.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Folders"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceFolders Folders
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceFolders>(this, "Folders", NetOffice.OfficeApi.SharedWorkspaceFolders.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Links"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.SharedWorkspaceLinks Links
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OfficeApi.SharedWorkspaceLinks>(this, "Links", NetOffice.OfficeApi.SharedWorkspaceLinks.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Parent"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.URL"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string URL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "URL");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Connected"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public bool Connected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Connected");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.LastRefreshed"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public object LastRefreshed
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LastRefreshed");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.SourceURL"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string SourceURL
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SourceURL");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SourceURL", value);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Refresh"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.CreateNew"/> </remarks>
		/// <param name="uRL">optional object uRL</param>
		/// <param name="name">optional object name</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void CreateNew(object uRL, object name)
		{
			 Factory.ExecuteMethod(this, "CreateNew", uRL, name);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.CreateNew"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void CreateNew()
		{
			 Factory.ExecuteMethod(this, "CreateNew");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.CreateNew"/> </remarks>
		/// <param name="uRL">optional object uRL</param>
		[CustomMethod]
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void CreateNew(object uRL)
		{
			 Factory.ExecuteMethod(this, "CreateNew", uRL);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Delete"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.RemoveDocument"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void RemoveDocument()
		{
			 Factory.ExecuteMethod(this, "RemoveDocument");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Office.SharedWorkspace.Disconnect"/> </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void Disconnect()
		{
			 Factory.ExecuteMethod(this, "Disconnect");
		}

		#endregion

		#pragma warning restore
	}
}
