using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// IDataSourceControl
	/// </summary>
	[SyntaxBypass]
 	public class IDataSourceControl_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDataSourceControl_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDataSourceControl_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		/// <param name="dataMember">optional object dataMember</param>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.ProviderType get_ProviderType(object dataMember)
		{
			return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType", dataMember);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ProviderType
		/// </summary>
		/// <param name="dataMember">optional object dataMember</param>
		[SupportByVersion("OWC10", 1), Redirect("get_ProviderType")]
		public NetOffice.OWC10Api.Enums.ProviderType ProviderType(object dataMember)
		{
			return get_ProviderType(dataMember);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface IDataSourceControl 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IDataSourceControl : IDataSourceControl_
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
                    _type = typeof(IDataSourceControl);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDataSourceControl(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDataSourceControl(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataSourceControl(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionString");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionString", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CurrentDirectory
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "CurrentDirectory");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool UseRemoteProvider
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseRemoteProvider");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseRemoteProvider", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", NetOffice.ADODBApi.Connection.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DataEntry
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DataEntry");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DataEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MaxRecords
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxRecords");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxRecords", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset DefaultRecordset
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Recordset>(this, "DefaultRecordset", NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsources SchemaRowsources
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRowsources>(this, "SchemaRowsources", NetOffice.OWC10Api.SchemaRowsources.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationships SchemaRelationships
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRelationships>(this, "SchemaRelationships", NetOffice.OWC10Api.SchemaRelationships.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PageRowsources PageRowsources
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageRowsources>(this, "PageRowsources", NetOffice.OWC10Api.PageRowsources.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDefs RecordsetDefs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDefs>(this, "RecordsetDefs", NetOffice.OWC10Api.RecordsetDefs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.RecordsetDefs RootRecordsetDefs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDefs>(this, "RootRecordsetDefs", NetOffice.OWC10Api.RecordsetDefs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PivotDefs
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "PivotDefs");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DefaultRecordsetName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultRecordsetName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DefaultRecordsetName", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string XMLData
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.GroupLevels GroupLevels
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.GroupLevels>(this, "GroupLevels", NetOffice.OWC10Api.GroupLevels.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("OWC10", 1), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Constants
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Constants");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.ElementExtensions ElementExtensions
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ElementExtensions>(this, "ElementExtensions", NetOffice.OWC10Api.ElementExtensions.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsNew
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsNew");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsNew", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum RecordsetType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum>(this, "RecordsetType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "RecordsetType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.AllPageFields AllPageFields
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.AllPageFields>(this, "AllPageFields", NetOffice.OWC10Api.AllPageFields.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Section CurrentSection
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Section>(this, "CurrentSection", NetOffice.OWC10Api.Section.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.ProviderType ProviderType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.AllGroupingDefs AllGroupingDefs
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.AllGroupingDefs>(this, "AllGroupingDefs", NetOffice.OWC10Api.AllGroupingDefs.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool DisplayAlerts
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "DisplayAlerts");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "DisplayAlerts", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.DataPages DataPages
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.DataPages>(this, "DataPages", NetOffice.OWC10Api.DataPages.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GridX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridX", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 GridY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "GridY");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "GridY", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LoadError
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LoadError");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DefaultControlTypeEnum DefaultControlType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DefaultControlTypeEnum>(this, "DefaultControlType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DefaultControlType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsDirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Busy
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Busy");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string Version
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Version");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 MajorVersion
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MajorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string MinorVersion
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "MinorVersion");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string BuildNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "BuildNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string RevisionNumber
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "RevisionNumber");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsDataModelDirty
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "IsDataModelDirty");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "IsDataModelDirty", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscOfflineTypeEnum OfflineType
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscOfflineTypeEnum>(this, "OfflineType");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "OfflineType", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string OfflinePublication
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OfflinePublication");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OfflinePublication", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool Offline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Offline");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string OfflineSource
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "OfflineSource");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "OfflineSource", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscXMLLocationEnum XMLLocation
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscXMLLocationEnum>(this, "XMLLocation");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "XMLLocation", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public bool UseXMLData
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "UseXMLData");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "UseXMLData", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string XMLDataTarget
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "XMLDataTarget");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "XMLDataTarget", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public string ConnectionFile
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionFile");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ConnectionFile", value);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DefaultRecordsetDefName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DefaultRecordsetDefName");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ConnectionStringFullPath
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ConnectionStringFullPath");
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.SchemaDiagrams SchemaDiagrams
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaDiagrams>(this, "SchemaDiagrams", NetOffice.OWC10Api.SchemaDiagrams.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DBNSOwnerName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "DBNSOwnerName");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordsetName">string recordsetName</param>
		/// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
		/// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum FetchType = 2</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption, object fetchType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType, recordsetName, executeOption, fetchType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordsetName">string recordsetName</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType, recordsetName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordsetName">string recordsetName</param>
		/// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType, recordsetName, executeOption);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dataAssistant">object dataAssistant</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SetDataAssistant(object dataAssistant)
		{
			 Factory.ExecuteMethod(this, "SetDataAssistant", dataAssistant);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="advise">object advise</param>
		/// <param name="sinkName">string sinkName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DesignAdvise(object advise, string sinkName)
		{
			 Factory.ExecuteMethod(this, "DesignAdvise", advise, sinkName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sinkName">string sinkName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DesignUnAdvise(string sinkName)
		{
			 Factory.ExecuteMethod(this, "DesignUnAdvise", sinkName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pUnknownDropGoo">object pUnknownDropGoo</param>
		/// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="pageRowsource">string pageRowsource</param>
		/// <param name="schemaRelationship">string schemaRelationship</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ProcessDrop(object pUnknownDropGoo, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
		{
			 Factory.ExecuteMethod(this, "ProcessDrop", new object[]{ pUnknownDropGoo, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="rowsources">object rowsources</param>
		/// <param name="relationships">object relationships</param>
		/// <param name="fields">object fields</param>
		/// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="pageRowsource">string pageRowsource</param>
		/// <param name="schemaRelationship">string schemaRelationship</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ScriptDrop(object rowsources, object relationships, object fields, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
		{
			 Factory.ExecuteMethod(this, "ScriptDrop", new object[]{ rowsources, relationships, fields, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="element">object element</param>
		[SupportByVersion("OWC10", 1)]
		public NetOffice.OWC10Api.Section GetContainingSection(object element)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Section>(this, "GetContainingSection", NetOffice.OWC10Api.Section.LateBindingApiWrapperType, element);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="rowsources">object rowsources</param>
		/// <param name="relationships">object relationships</param>
		/// <param name="fields">object fields</param>
		/// <param name="recordsetDef">string recordsetDef</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="dropRowsource">string dropRowsource</param>
		/// <param name="rowsourcesOut">object rowsourcesOut</param>
		/// <param name="relationshipsOut">object relationshipsOut</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ScriptValidate(object rowsources, object relationships, object fields, string recordsetDef, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,true,true,true);
			dropRowsource = string.Empty;
			rowsourcesOut = null;
			relationshipsOut = null;
			object[] paramsArray = Invoker.ValidateParamsArray(rowsources, relationships, fields, recordsetDef, dl, dt, dropRowsource, rowsourcesOut, relationshipsOut);
			Invoker.Method(this, "ScriptValidate", paramsArray, modifiers);
			dropRowsource = paramsArray[6] as string;
			rowsourcesOut = (object)paramsArray[7];
			relationshipsOut = (object)paramsArray[8];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="unknownDropGoo">object unknownDropGoo</param>
		/// <param name="recordSetDefName">string recordSetDefName</param>
		/// <param name="location">NetOffice.OWC10Api.Enums.DscDropLocationEnum location</param>
		/// <param name="type">NetOffice.OWC10Api.Enums.DscDropTypeEnum type</param>
		/// <param name="dropRowsource">string dropRowsource</param>
		/// <param name="rowsourcesOut">object rowsourcesOut</param>
		/// <param name="relationshipsOut">object relationshipsOut</param>
		/// <param name="numberOfDrops">Int32 numberOfDrops</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ValidateDrop(object unknownDropGoo, string recordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum location, NetOffice.OWC10Api.Enums.DscDropTypeEnum type, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut, out Int32 numberOfDrops)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,true,true,true);
			dropRowsource = string.Empty;
			rowsourcesOut = null;
			relationshipsOut = null;
			numberOfDrops = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(unknownDropGoo, recordSetDefName, location, type, dropRowsource, rowsourcesOut, relationshipsOut, numberOfDrops);
			Invoker.Method(this, "ValidateDrop", paramsArray, modifiers);
			dropRowsource = paramsArray[4] as string;
			rowsourcesOut = (object)paramsArray[5];
			relationshipsOut = (object)paramsArray[6];
			numberOfDrops = (Int32)paramsArray[7];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="hyperlink">object hyperlink</param>
		/// <param name="part">NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public string HyperlinkPart(object hyperlink, NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part)
		{
			return Factory.ExecuteStringMethodGet(this, "HyperlinkPart", hyperlink, part);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SchemaRefresh()
		{
			 Factory.ExecuteMethod(this, "SchemaRefresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="oldID">string oldID</param>
		/// <param name="newID">string newID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void UpdateElementID(string oldID, string newID)
		{
			 Factory.ExecuteMethod(this, "UpdateElementID", oldID, newID);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void Reset()
		{
			 Factory.ExecuteMethod(this, "Reset");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public string getDataMemberName(Int32 lIndex)
		{
			return Factory.ExecuteStringMethodGet(this, "getDataMemberName", lIndex);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public Int32 getDataMemberCount()
		{
			return Factory.ExecuteInt32MethodGet(this, "getDataMemberCount");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectionElement">object sectionElement</param>
		/// <param name="recordSource">string recordSource</param>
		/// <param name="sectionType">NetOffice.OWC10Api.Enums.SectTypeEnum sectionType</param>
		/// <param name="groupLevel">NetOffice.OWC10Api.GroupLevel groupLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void GetSectionInfo(object sectionElement, out string recordSource, out NetOffice.OWC10Api.Enums.SectTypeEnum sectionType, out NetOffice.OWC10Api.GroupLevel groupLevel)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true);
			recordSource = string.Empty;
			sectionType = 0;
			groupLevel = null;
			object[] paramsArray = Invoker.ValidateParamsArray(sectionElement, recordSource, sectionType, groupLevel);
			Invoker.Method(this, "GetSectionInfo", paramsArray, modifiers);
			recordSource = paramsArray[1] as string;
			sectionType = (NetOffice.OWC10Api.Enums.SectTypeEnum)paramsArray[2];
            if (paramsArray[3] is MarshalByRefObject)
                groupLevel = new NetOffice.OWC10Api.GroupLevel(this, paramsArray[3]);
            else
                groupLevel = null;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DeleteRecordSourceIfUnused(string recordSource)
		{
			 Factory.ExecuteMethod(this, "DeleteRecordSourceIfUnused", recordSource);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordSource">string recordSource</param>
		/// <param name="pageField">string pageField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DeletePageFieldIfUnused(string recordSource, string pageField)
		{
			 Factory.ExecuteMethod(this, "DeletePageFieldIfUnused", recordSource, pageField);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="bstrRecordset">string bstrRecordset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void ResetRecordset(string bstrRecordset)
		{
			 Factory.ExecuteMethod(this, "ResetRecordset", bstrRecordset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="exportType">NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType</param>
		/// <param name="connectString">string connectString</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void GetExportableConnectString(NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType, out string connectString)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			connectString = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(exportType, connectString);
			Invoker.Method(this, "GetExportableConnectString", paramsArray, modifiers);
			connectString = paramsArray[1] as string;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="eEncoding">optional NetOffice.OWC10Api.Enums.DscEncodingEnum eEncoding = 0</param>
		[SupportByVersion("OWC10", 1)]
		public void ExportXML(object eEncoding)
		{
			 Factory.ExecuteMethod(this, "ExportXML", eEncoding);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void ExportXML()
		{
			 Factory.ExecuteMethod(this, "ExportXML");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="recordsetName">string recordsetName</param>
		/// <param name="recordset">NetOffice.ADODBApi.Recordset recordset</param>
		[SupportByVersion("OWC10", 1)]
		public void SetRootRecordset(string recordsetName, NetOffice.ADODBApi.Recordset recordset)
		{
			 Factory.ExecuteMethod(this, "SetRootRecordset", recordsetName, recordset);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="onlineServer">string onlineServer</param>
		/// <param name="onlineDatabase">string onlineDatabase</param>
		/// <param name="offlineServer">string offlineServer</param>
		/// <param name="offlineDatabase">string offlineDatabase</param>
		[SupportByVersion("OWC10", 1)]
		public void GetOfflineDisplayInfo(out string onlineServer, out string onlineDatabase, out string offlineServer, out string offlineDatabase)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			onlineServer = string.Empty;
			onlineDatabase = string.Empty;
			offlineServer = string.Empty;
			offlineDatabase = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(onlineServer, onlineDatabase, offlineServer, offlineDatabase);
			Invoker.Method(this, "GetOfflineDisplayInfo", paramsArray, modifiers);
			onlineServer = paramsArray[0] as string;
			onlineDatabase = paramsArray[1] as string;
			offlineServer = paramsArray[2] as string;
			offlineDatabase = paramsArray[3] as string;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="refreshType">optional NetOffice.OWC10Api.Enums.RefreshType RefreshType = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void Refresh(object refreshType)
		{
			 Factory.ExecuteMethod(this, "Refresh", refreshType);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public void Refresh()
		{
			 Factory.ExecuteMethod(this, "Refresh");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="fChild">Int32 fChild</param>
		/// <param name="ppGrouplevel">NetOffice.OWC10Api.GroupLevel ppGrouplevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void FindRelatedGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, Int32 fChild, out NetOffice.OWC10Api.GroupLevel ppGrouplevel)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppGrouplevel = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pGroupLevel, fChild, ppGrouplevel);
			Invoker.Method(this, "FindRelatedGroupLevel", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppGrouplevel = new NetOffice.OWC10Api.GroupLevel(this, paramsArray[2]);
            else
                ppGrouplevel = null;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="notification">NetOffice.OWC10Api.Enums.NotificationType notification</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void DllNotification(NetOffice.OWC10Api.Enums.NotificationType notification)
		{
			 Factory.ExecuteMethod(this, "DllNotification", notification);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="suspend">bool suspend</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SuspendUndo(bool suspend)
		{
			 Factory.ExecuteMethod(this, "SuspendUndo", suspend);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void UpdateFocus()
		{
			 Factory.ExecuteMethod(this, "UpdateFocus");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="connectionString">string connectionString</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public bool IsValidDAPProvider(string connectionString)
		{
			return Factory.ExecuteBoolMethodGet(this, "IsValidDAPProvider", connectionString);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		/// <param name="triangulationPrecision">optional object triangulationPrecision</param>
		[SupportByVersion("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
		{
			return Factory.ExecuteDoubleMethodGet(this, "EuroConvert", new object[]{ number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision });
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
		{
			return Factory.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="number">Double number</param>
		/// <param name="sourceCurrency">string sourceCurrency</param>
		/// <param name="targetCurrency">string targetCurrency</param>
		/// <param name="fullPrecision">optional object fullPrecision</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
		{
			return Factory.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency, fullPrecision);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public String[] GetDAPProviders()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetDAPProviders", paramsArray);
			return (String[])returnItem;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="synchronizing">bool synchronizing</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SetSynchronizing(bool synchronizing)
		{
			 Factory.ExecuteMethod(this, "SetSynchronizing", synchronizing);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="displayError">bool displayError</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SetDisplayError(bool displayError)
		{
			 Factory.ExecuteMethod(this, "SetDisplayError", displayError);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="suspend">bool suspend</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SuspendXMLReExecute(bool suspend)
		{
			 Factory.ExecuteMethod(this, "SuspendXMLReExecute", suspend);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="firePropChange">bool firePropChange</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void SetFirePropChange(bool firePropChange)
		{
			 Factory.ExecuteMethod(this, "SetFirePropChange", firePropChange);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="value">object value</param>
		/// <param name="valueIfNull">optional object valueIfNull</param>
		[SupportByVersion("OWC10", 1)]
		public object Nz(object value, object valueIfNull)
		{
			return Factory.ExecuteVariantMethodGet(this, "Nz", value, valueIfNull);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="value">object value</param>
		[CustomMethod]
		[SupportByVersion("OWC10", 1)]
		public object Nz(object value)
		{
			return Factory.ExecuteVariantMethodGet(this, "Nz", value);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public void RefreshJetCache()
		{
			 Factory.ExecuteMethod(this, "RefreshJetCache");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("OWC10", 1)]
		public void AutoRefreshOfflineSource()
		{
			 Factory.ExecuteMethod(this, "AutoRefreshOfflineSource");
		}

		#endregion

		#pragma warning restore
	}
}
