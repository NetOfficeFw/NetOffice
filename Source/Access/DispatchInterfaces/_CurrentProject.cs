using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface _CurrentProject 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CurrentProject : COMObject
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
                    _type = typeof(_CurrentProject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public _CurrentProject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public _CurrentProject(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AllForms"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AllForms AllForms
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllForms>(this, "AllForms", NetOffice.AccessApi.AllForms.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AllReports"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AllReports AllReports
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllReports>(this, "AllReports", NetOffice.AccessApi.AllReports.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AllMacros"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AllMacros AllMacros
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllMacros>(this, "AllMacros", NetOffice.AccessApi.AllMacros.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AllModules"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AllModules AllModules
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllModules>(this, "AllModules", NetOffice.AccessApi.AllModules.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AllDataAccessPages AllDataAccessPages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllDataAccessPages>(this, "AllDataAccessPages", NetOffice.AccessApi.AllDataAccessPages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.ProjectType"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcProjectType ProjectType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcProjectType>(this, "ProjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.BaseConnectionString"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string BaseConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BaseConnectionString");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.IsConnected"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool IsConnected
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Name"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Path"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string Path
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.FullName"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Connection"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Properties"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.AccessObjectProperties Properties
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AccessObjectProperties>(this, "Properties", NetOffice.AccessApi.AccessObjectProperties.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Application"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", NetOffice.AccessApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Parent"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.RemovePersonalInformation"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public bool RemovePersonalInformation
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.FileFormat"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcFileFormat FileFormat
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFileFormat>(this, "FileFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AccessConnection"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.ADODBApi.Connection AccessConnection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "AccessConnection", NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.ImportExportSpecifications"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public NetOffice.AccessApi.ImportExportSpecifications ImportExportSpecifications
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.ImportExportSpecifications>(this, "ImportExportSpecifications", NetOffice.AccessApi.ImportExportSpecifications.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// Returns True if the current project was created in Access 2013 and onwards and False if the current project was created prior to Access 2013. Read-only Boolean.
		/// SupportByVersion Access 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/access/concepts/miscellaneous/currentproject-issqlbackend-property-access"/> </remarks>
		[SupportByVersion("Access", 16)]
		public bool IsSQLBackend
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsSQLBackend");
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.IsTrusted"/> </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public bool IsTrusted
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsTrusted");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.WebSite"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public string WebSite
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WebSite");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.IsWeb"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool IsWeb
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsWeb");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.Resources"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public NetOffice.AccessApi.SharedResources Resources
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.SharedResources>(this, "Resources", NetOffice.AccessApi.SharedResources.LateBindingApiWrapperType);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.OpenConnection"/> </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void OpenConnection(object baseConnectionString, object userID, object password)
		{
			 Factory.ExecuteMethod(this, "OpenConnection", baseConnectionString, userID, password);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.OpenConnection"/> </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void OpenConnection()
		{
			 Factory.ExecuteMethod(this, "OpenConnection");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.OpenConnection"/> </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void OpenConnection(object baseConnectionString)
		{
			 Factory.ExecuteMethod(this, "OpenConnection", baseConnectionString);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.OpenConnection"/> </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		/// <param name="userID">optional object userID</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void OpenConnection(object baseConnectionString, object userID)
		{
			 Factory.ExecuteMethod(this, "OpenConnection", baseConnectionString, userID);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.CloseConnection"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public void CloseConnection()
		{
			 Factory.ExecuteMethod(this, "CloseConnection");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.UpdateDependencyInfo"/> </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public void UpdateDependencyInfo()
		{
			 Factory.ExecuteMethod(this, "UpdateDependencyInfo");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsMemberSafe(Int32 dispid)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.CurrentProject.AddSharedImage"/> </remarks>
		/// <param name="sharedImageName">string sharedImageName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public void AddSharedImage(string sharedImageName, string fileName)
		{
			 Factory.ExecuteMethod(this, "AddSharedImage", sharedImageName, fileName);
		}

		#endregion

		#pragma warning restore
	}
}
