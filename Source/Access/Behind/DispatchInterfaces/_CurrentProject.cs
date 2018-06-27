using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.AccessApi;

namespace NetOffice.AccessApi.Behind
{
	/// <summary>
	/// DispatchInterface _CurrentProject 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class _CurrentProject : COMObject, NetOffice.AccessApi._CurrentProject
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Contract Type
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
        public override Type ContractType
        {
            get
            {
                if(null == _contractType)
                    _contractType = typeof(NetOffice.AccessApi._CurrentProject);
                return _contractType;
            }
        }
        private static Type _contractType;


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

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public _CurrentProject() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193455.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AllForms AllForms
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllForms>(this, "AllForms", typeof(NetOffice.AccessApi.AllForms));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835381.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AllReports AllReports
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllReports>(this, "AllReports", typeof(NetOffice.AccessApi.AllReports));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195863.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AllMacros AllMacros
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllMacros>(this, "AllMacros", typeof(NetOffice.AccessApi.AllMacros));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192113.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AllModules AllModules
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllModules>(this, "AllModules", typeof(NetOffice.AccessApi.AllModules));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AllDataAccessPages AllDataAccessPages
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AllDataAccessPages>(this, "AllDataAccessPages", typeof(NetOffice.AccessApi.AllDataAccessPages));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822409.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcProjectType ProjectType
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcProjectType>(this, "ProjectType");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191980.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string BaseConnectionString
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BaseConnectionString");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff844855.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual bool IsConnected
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsConnected");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822433.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Name
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Name");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff191878.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string Path
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Path");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff193148.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual string FullName
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "FullName");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821478.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", typeof(NetOffice.ADODBApi.Connection));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837238.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.AccessObjectProperties Properties
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.AccessObjectProperties>(this, "Properties", typeof(NetOffice.AccessApi.AccessObjectProperties));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff194210.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.Application>(this, "Application", typeof(NetOffice.AccessApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821768.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835099.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual bool RemovePersonalInformation
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "RemovePersonalInformation");
			}
			set
			{
				InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "RemovePersonalInformation", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836251.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.AccessApi.Enums.AcFileFormat FileFormat
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcFileFormat>(this, "FileFormat");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823021.aspx </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public virtual NetOffice.ADODBApi.Connection AccessConnection
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "AccessConnection", typeof(NetOffice.ADODBApi.Connection));
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff822096.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual NetOffice.AccessApi.ImportExportSpecifications ImportExportSpecifications
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.ImportExportSpecifications>(this, "ImportExportSpecifications", typeof(NetOffice.AccessApi.ImportExportSpecifications));
			}
		}

		/// <summary>
		/// SupportByVersion Access 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823046.aspx </remarks>
		[SupportByVersion("Access", 12,14,15,16)]
		public virtual bool IsTrusted
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsTrusted");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821472.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual string WebSite
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "WebSite");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835346.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual bool IsWeb
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsWeb");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192241.aspx </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public virtual NetOffice.AccessApi.SharedResources Resources
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.AccessApi.SharedResources>(this, "Resources", typeof(NetOffice.AccessApi.SharedResources));
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192495.aspx </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		/// <param name="userID">optional object userID</param>
		/// <param name="password">optional object password</param>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenConnection(object baseConnectionString, object userID, object password)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenConnection", baseConnectionString, userID, password);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192495.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenConnection()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenConnection");
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192495.aspx </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenConnection(object baseConnectionString)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenConnection", baseConnectionString);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff192495.aspx </remarks>
		/// <param name="baseConnectionString">optional object baseConnectionString</param>
		/// <param name="userID">optional object userID</param>
		[CustomMethod]
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void OpenConnection(object baseConnectionString, object userID)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "OpenConnection", baseConnectionString, userID);
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836645.aspx </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public virtual void CloseConnection()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "CloseConnection");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff197401.aspx </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual void UpdateDependencyInfo()
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateDependencyInfo");
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="dispid">Int32 dispid</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Access", 11,12,14,15,16)]
		public virtual bool IsMemberSafe(Int32 dispid)
		{
			return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsMemberSafe", dispid);
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823140.aspx </remarks>
		/// <param name="sharedImageName">string sharedImageName</param>
		/// <param name="fileName">string fileName</param>
		[SupportByVersion("Access", 14,15,16)]
		public virtual void AddSharedImage(string sharedImageName, string fileName)
		{
			 InvokerService.InvokeInternal.ExecuteMethod(this, "AddSharedImage", sharedImageName, fileName);
		}

		#endregion

		#pragma warning restore
	}
}


