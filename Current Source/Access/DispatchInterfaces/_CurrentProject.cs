using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using LateBindingApi.Core;
namespace NetOffice.AccessApi
{
	///<summary>
	/// DispatchInterface _CurrentProject 
	/// SupportByVersion Access, 9,10,11,12,14
	///</summary>
	[SupportByVersionAttribute("Access", 9,10,11,12,14)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class _CurrentProject : COMObject
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
                    _type = typeof(_CurrentProject);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public _CurrentProject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AllForms AllForms
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllForms", paramsArray);
				NetOffice.AccessApi.AllForms newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AllForms.LateBindingApiWrapperType) as NetOffice.AccessApi.AllForms;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AllReports AllReports
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllReports", paramsArray);
				NetOffice.AccessApi.AllReports newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AllReports.LateBindingApiWrapperType) as NetOffice.AccessApi.AllReports;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AllMacros AllMacros
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllMacros", paramsArray);
				NetOffice.AccessApi.AllMacros newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AllMacros.LateBindingApiWrapperType) as NetOffice.AccessApi.AllMacros;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AllModules AllModules
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllModules", paramsArray);
				NetOffice.AccessApi.AllModules newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AllModules.LateBindingApiWrapperType) as NetOffice.AccessApi.AllModules;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AllDataAccessPages AllDataAccessPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllDataAccessPages", paramsArray);
				NetOffice.AccessApi.AllDataAccessPages newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AllDataAccessPages.LateBindingApiWrapperType) as NetOffice.AccessApi.AllDataAccessPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.Enums.AcProjectType ProjectType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProjectType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.AccessApi.Enums.AcProjectType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public string BaseConnectionString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BaseConnectionString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public bool IsConnected
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsConnected", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public string Name
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Name", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public string Path
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Path", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public string FullName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FullName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connection", paramsArray);
				NetOffice.ADODBApi.Connection newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Connection.LateBindingApiWrapperType) as NetOffice.ADODBApi.Connection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.AccessObjectProperties Properties
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Properties", paramsArray);
				NetOffice.AccessApi.AccessObjectProperties newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.AccessObjectProperties.LateBindingApiWrapperType) as NetOffice.AccessApi.AccessObjectProperties;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public NetOffice.AccessApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.AccessApi.Application newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.Application.LateBindingApiWrapperType) as NetOffice.AccessApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = LateBindingApi.Core.Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14)]
		public bool RemovePersonalInformation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RemovePersonalInformation", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RemovePersonalInformation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14)]
		public NetOffice.AccessApi.Enums.AcFileFormat FileFormat
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "FileFormat", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.AccessApi.Enums.AcFileFormat)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 10,11,12,14)]
		public NetOffice.ADODBApi.Connection AccessConnection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AccessConnection", paramsArray);
				NetOffice.ADODBApi.Connection newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Connection.LateBindingApiWrapperType) as NetOffice.ADODBApi.Connection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14)]
		public NetOffice.AccessApi.ImportExportSpecifications ImportExportSpecifications
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ImportExportSpecifications", paramsArray);
				NetOffice.AccessApi.ImportExportSpecifications newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.ImportExportSpecifications.LateBindingApiWrapperType) as NetOffice.AccessApi.ImportExportSpecifications;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 12,14)]
		public bool IsTrusted
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsTrusted", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 14)]
		public string WebSite
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "WebSite", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 14)]
		public bool IsWeb
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsWeb", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion Access 14
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Access", 14)]
		public NetOffice.AccessApi.SharedResources Resources
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Resources", paramsArray);
				NetOffice.AccessApi.SharedResources newObject = LateBindingApi.Core.Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.AccessApi.SharedResources.LateBindingApiWrapperType) as NetOffice.AccessApi.SharedResources;
				return newObject;
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="baseConnectionString">optional object BaseConnectionString</param>
		/// <param name="userID">optional object UserID</param>
		/// <param name="password">optional object Password</param>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void OpenConnection(object baseConnectionString, object userID, object password)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(baseConnectionString, userID, password);
			Invoker.Method(this, "OpenConnection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void OpenConnection()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "OpenConnection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="baseConnectionString">optional object BaseConnectionString</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void OpenConnection(object baseConnectionString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(baseConnectionString);
			Invoker.Method(this, "OpenConnection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		/// <param name="baseConnectionString">optional object BaseConnectionString</param>
		/// <param name="userID">optional object UserID</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void OpenConnection(object baseConnectionString, object userID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(baseConnectionString, userID);
			Invoker.Method(this, "OpenConnection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Access", 9,10,11,12,14)]
		public void CloseConnection()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "CloseConnection", paramsArray);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14
		/// </summary>
		[SupportByVersionAttribute("Access", 11,12,14)]
		public void UpdateDependencyInfo()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateDependencyInfo", paramsArray);
		}

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

		/// <summary>
		/// SupportByVersion Access 14
		/// </summary>
		/// <param name="sharedImageName">string SharedImageName</param>
		/// <param name="fileName">string FileName</param>
		[SupportByVersionAttribute("Access", 14)]
		public void AddSharedImage(string sharedImageName, string fileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sharedImageName, fileName);
			Invoker.Method(this, "AddSharedImage", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}