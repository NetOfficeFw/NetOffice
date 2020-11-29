using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.AccessApi
{
	/// <summary>
	/// DispatchInterface AccessObject 
	/// SupportByVersion Access, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject"/> </remarks>
	[SupportByVersion("Access", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface)]
 	public class AccessObject : COMObject
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
                    _type = typeof(AccessObject);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public AccessObject(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public AccessObject(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public AccessObject(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.Parent"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string _Name
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "_Name");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.Name"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.Type"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public Int32 Type
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Type");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 Attributes
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Attributes");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.Properties"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.IsLoaded"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public bool IsLoaded
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsLoaded");
			}
		}

		/// <summary>
		/// SupportByVersion Access 9, 10, 11, 12, 14, 15, 16
		/// Get/Set
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.FullName"/> </remarks>
		[SupportByVersion("Access", 9,10,11,12,14,15,16)]
		public string FullName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FullName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FullName", value);
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.DateCreated"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public DateTime DateCreated
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "DateCreated");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.DateModified"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public DateTime DateModified
		{
			get
			{
				return Factory.ExecuteDateTimePropertyGet(this, "DateModified");
			}
		}

		/// <summary>
		/// SupportByVersion Access 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.CurrentView"/> </remarks>
		[SupportByVersion("Access", 10,11,12,14,15,16)]
		public NetOffice.AccessApi.Enums.AcCurrentView CurrentView
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.AccessApi.Enums.AcCurrentView>(this, "CurrentView");
			}
		}

		/// <summary>
		/// SupportByVersion Access 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.IsWeb"/> </remarks>
		[SupportByVersion("Access", 14,15,16)]
		public bool IsWeb
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsWeb");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.IsDependentUpon"/> </remarks>
		/// <param name="objectType">NetOffice.AccessApi.Enums.AcObjectType objectType</param>
		/// <param name="objectName">string objectName</param>
		[SupportByVersion("Access", 11,12,14,15,16)]
		public bool IsDependentUpon(NetOffice.AccessApi.Enums.AcObjectType objectType, string objectName)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsDependentUpon", objectType, objectName);
		}

		/// <summary>
		/// SupportByVersion Access 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Access.AccessObject.GetDependencyInfo"/> </remarks>
		[SupportByVersion("Access", 11,12,14,15,16)]
		[BaseResult]
		public NetOffice.AccessApi._DependencyInfo GetDependencyInfo()
		{
			return Factory.ExecuteBaseReferenceMethodGet<NetOffice.AccessApi._DependencyInfo>(this, "GetDependencyInfo");
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

		#endregion

		#pragma warning restore
	}
}
