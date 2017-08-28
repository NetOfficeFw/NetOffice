using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OfficeApi
{
	/// <summary>
	/// DispatchInterface Sync 
	/// SupportByVersion Office, 11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860602.aspx </remarks>
	[SupportByVersion("Office", 11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class Sync : _IMsoDispObj
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
                    _type = typeof(Sync);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public Sync(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public Sync(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Sync(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865564.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoSyncStatusType Status
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSyncStatusType>(this, "Status");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff865364.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public string WorkspaceLastChangedBy
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "WorkspaceLastChangedBy");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864917.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public object LastSyncTime
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "LastSyncTime");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff862150.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public NetOffice.OfficeApi.Enums.MsoSyncErrorType ErrorType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OfficeApi.Enums.MsoSyncErrorType>(this, "ErrorType");
			}
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860559.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16), ProxyResult]
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
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff863651.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void GetUpdate()
		{
			 Factory.ExecuteMethod(this, "GetUpdate");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860754.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void PutUpdate()
		{
			 Factory.ExecuteMethod(this, "PutUpdate");
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff860783.aspx </remarks>
		/// <param name="syncVersionType">NetOffice.OfficeApi.Enums.MsoSyncVersionType syncVersionType</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void OpenVersion(NetOffice.OfficeApi.Enums.MsoSyncVersionType syncVersionType)
		{
			 Factory.ExecuteMethod(this, "OpenVersion", syncVersionType);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff864675.aspx </remarks>
		/// <param name="syncConflictResolution">NetOffice.OfficeApi.Enums.MsoSyncConflictResolutionType syncConflictResolution</param>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void ResolveConflict(NetOffice.OfficeApi.Enums.MsoSyncConflictResolutionType syncConflictResolution)
		{
			 Factory.ExecuteMethod(this, "ResolveConflict", syncConflictResolution);
		}

		/// <summary>
		/// SupportByVersion Office 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff861422.aspx </remarks>
		[SupportByVersion("Office", 11,12,14,15,16)]
		public void Unsuspend()
		{
			 Factory.ExecuteMethod(this, "Unsuspend");
		}

		#endregion

		#pragma warning restore
	}
}
