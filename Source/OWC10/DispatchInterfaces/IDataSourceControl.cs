using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using System.ComponentModel;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// IDataSourceControl
	///</summary>
	public class IDataSourceControl_ : COMObject
	{
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// <param name="dataMember">optional object DataMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.ProviderType get_ProviderType(object dataMember)
		{		
			object[] paramsArray = Invoker.ValidateParamsArray(dataMember);
			object returnItem = Invoker.PropertyGet(this, "ProviderType", paramsArray);
			int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
			return (NetOffice.OWC10Api.Enums.ProviderType)intReturnItem;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Alias for get_ProviderType
		/// </summary>
		/// <param name="dataMember">optional object DataMember</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.ProviderType ProviderType(object dataMember)
		{
			return get_ProviderType(dataMember);
		}

		#endregion

		#region Methods

		#endregion

	}

	///<summary>
	/// DispatchInterface IDataSourceControl 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsDispatchInterface)]
	public class IDataSourceControl : IDataSourceControl_
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
                    _type = typeof(IDataSourceControl);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		[SupportByVersionAttribute("OWC10", 1)]
		public string ConnectionString
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionString", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectionString", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string CurrentDirectory
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentDirectory", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool UseRemoteProvider
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseRemoteProvider", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseRemoteProvider", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.ADODBApi.Connection Connection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Connection", paramsArray);
				NetOffice.ADODBApi.Connection newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Connection.LateBindingApiWrapperType) as NetOffice.ADODBApi.Connection;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool DataEntry
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataEntry", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DataEntry", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 MaxRecords
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MaxRecords", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MaxRecords", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset DefaultRecordset
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultRecordset", paramsArray);
				NetOffice.ADODBApi.Recordset newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType) as NetOffice.ADODBApi.Recordset;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRowsources SchemaRowsources
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SchemaRowsources", paramsArray);
				NetOffice.OWC10Api.SchemaRowsources newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.SchemaRowsources.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaRowsources;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.SchemaRelationships SchemaRelationships
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SchemaRelationships", paramsArray);
				NetOffice.OWC10Api.SchemaRelationships newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.SchemaRelationships.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaRelationships;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.PageRowsources PageRowsources
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PageRowsources", paramsArray);
				NetOffice.OWC10Api.PageRowsources newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.PageRowsources.LateBindingApiWrapperType) as NetOffice.OWC10Api.PageRowsources;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.RecordsetDefs RecordsetDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordsetDefs", paramsArray);
				NetOffice.OWC10Api.RecordsetDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.RecordsetDefs.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.RecordsetDefs RootRecordsetDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RootRecordsetDefs", paramsArray);
				NetOffice.OWC10Api.RecordsetDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.RecordsetDefs.LateBindingApiWrapperType) as NetOffice.OWC10Api.RecordsetDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object PivotDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "PivotDefs", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DefaultRecordsetName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultRecordsetName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultRecordsetName", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string XMLData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLData", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.GroupLevels GroupLevels
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GroupLevels", paramsArray);
				NetOffice.OWC10Api.GroupLevels newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.GroupLevels.LateBindingApiWrapperType) as NetOffice.OWC10Api.GroupLevels;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Constants
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Constants", paramsArray);
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.ElementExtensions ElementExtensions
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ElementExtensions", paramsArray);
				NetOffice.OWC10Api.ElementExtensions newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.ElementExtensions.LateBindingApiWrapperType) as NetOffice.OWC10Api.ElementExtensions;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsNew
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsNew", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsNew", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum RecordsetType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RecordsetType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "RecordsetType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.AllPageFields AllPageFields
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllPageFields", paramsArray);
				NetOffice.OWC10Api.AllPageFields newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.AllPageFields.LateBindingApiWrapperType) as NetOffice.OWC10Api.AllPageFields;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section CurrentSection
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "CurrentSection", paramsArray);
				NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.Enums.ProviderType ProviderType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ProviderType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.ProviderType)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.AllGroupingDefs AllGroupingDefs
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "AllGroupingDefs", paramsArray);
				NetOffice.OWC10Api.AllGroupingDefs newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.AllGroupingDefs.LateBindingApiWrapperType) as NetOffice.OWC10Api.AllGroupingDefs;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool DisplayAlerts
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DisplayAlerts", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DisplayAlerts", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.DataPages DataPages
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DataPages", paramsArray);
				NetOffice.OWC10Api.DataPages newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.DataPages.LateBindingApiWrapperType) as NetOffice.OWC10Api.DataPages;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GridX
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridX", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridX", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GridY
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "GridY", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "GridY", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LoadError
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "LoadError", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DefaultControlTypeEnum DefaultControlType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultControlType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.DefaultControlTypeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "DefaultControlType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsDirty
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDirty", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsDirty", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Busy
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Busy", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string Version
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Version", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 MajorVersion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MajorVersion", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string MinorVersion
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MinorVersion", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string BuildNumber
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "BuildNumber", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string RevisionNumber
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "RevisionNumber", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool IsDataModelDirty
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "IsDataModelDirty", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "IsDataModelDirty", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscOfflineTypeEnum OfflineType
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OfflineType", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.DscOfflineTypeEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OfflineType", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string OfflinePublication
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OfflinePublication", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OfflinePublication", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool Offline
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Offline", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string OfflineSource
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "OfflineSource", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "OfflineSource", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Enums.DscXMLLocationEnum XMLLocation
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLLocation", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.OWC10Api.Enums.DscXMLLocationEnum)intReturnItem;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLLocation", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public bool UseXMLData
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "UseXMLData", paramsArray);
				return NetRuntimeSystem.Convert.ToBoolean(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "UseXMLData", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string XMLDataTarget
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "XMLDataTarget", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "XMLDataTarget", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get/Set
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public string ConnectionFile
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionFile", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "ConnectionFile", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DefaultRecordsetDefName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DefaultRecordsetDefName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string ConnectionStringFullPath
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "ConnectionStringFullPath", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.OWC10Api.SchemaDiagrams SchemaDiagrams
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "SchemaDiagrams", paramsArray);
				NetOffice.OWC10Api.SchemaDiagrams newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.OWC10Api.SchemaDiagrams.LateBindingApiWrapperType) as NetOffice.OWC10Api.SchemaDiagrams;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// Get
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string DBNSOwnerName
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "DBNSOwnerName", paramsArray);
				return NetRuntimeSystem.Convert.ToString(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordsetName">string RecordsetName</param>
		/// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
		/// <param name="fetchType">optional NetOffice.OWC10Api.Enums.DscFetchTypeEnum FetchType = 2</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption, object fetchType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsetName, executeOption, fetchType);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi.Recordset newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType) as NetOffice.ADODBApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordsetName">string RecordsetName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsetName);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi.Recordset newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType) as NetOffice.ADODBApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordsetName">string RecordsetName</param>
		/// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsetName, executeOption);
			object returnItem = Invoker.MethodReturn(this, "Execute", paramsArray);
			NetOffice.ADODBApi.Recordset newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ADODBApi.Recordset.LateBindingApiWrapperType) as NetOffice.ADODBApi.Recordset;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dataAssistant">object DataAssistant</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetDataAssistant(object dataAssistant)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dataAssistant);
			Invoker.Method(this, "SetDataAssistant", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="advise">object Advise</param>
		/// <param name="sinkName">string SinkName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DesignAdvise(object advise, string sinkName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(advise, sinkName);
			Invoker.Method(this, "DesignAdvise", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="sinkName">string SinkName</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DesignUnAdvise(string sinkName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sinkName);
			Invoker.Method(this, "DesignUnAdvise", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pUnknownDropGoo">object pUnknownDropGoo</param>
		/// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="pageRowsource">string PageRowsource</param>
		/// <param name="schemaRelationship">string SchemaRelationship</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ProcessDrop(object pUnknownDropGoo, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pUnknownDropGoo, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship);
			Invoker.Method(this, "ProcessDrop", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="rowsources">object Rowsources</param>
		/// <param name="relationships">object Relationships</param>
		/// <param name="fields">object Fields</param>
		/// <param name="bstrRecordSetDefName">string bstrRecordSetDefName</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="pageRowsource">string PageRowsource</param>
		/// <param name="schemaRelationship">string SchemaRelationship</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ScriptDrop(object rowsources, object relationships, object fields, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(rowsources, relationships, fields, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship);
			Invoker.Method(this, "ScriptDrop", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="element">object Element</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public NetOffice.OWC10Api.Section GetContainingSection(object element)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(element);
			object returnItem = Invoker.MethodReturn(this, "GetContainingSection", paramsArray);
			NetOffice.OWC10Api.Section newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.OWC10Api.Section.LateBindingApiWrapperType) as NetOffice.OWC10Api.Section;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="rowsources">object Rowsources</param>
		/// <param name="relationships">object Relationships</param>
		/// <param name="fields">object Fields</param>
		/// <param name="recordsetDef">string RecordsetDef</param>
		/// <param name="dl">NetOffice.OWC10Api.Enums.DscDropLocationEnum dl</param>
		/// <param name="dt">NetOffice.OWC10Api.Enums.DscDropTypeEnum dt</param>
		/// <param name="dropRowsource">string DropRowsource</param>
		/// <param name="rowsourcesOut">object RowsourcesOut</param>
		/// <param name="relationshipsOut">object RelationshipsOut</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ScriptValidate(object rowsources, object relationships, object fields, string recordsetDef, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,false,false,true,true,true);
			dropRowsource = string.Empty;
			rowsourcesOut = null;
			relationshipsOut = null;
			object[] paramsArray = Invoker.ValidateParamsArray(rowsources, relationships, fields, recordsetDef, dl, dt, dropRowsource, rowsourcesOut, relationshipsOut);
			Invoker.Method(this, "ScriptValidate", paramsArray, modifiers);
			dropRowsource = (string)paramsArray[6];
			rowsourcesOut = (object)paramsArray[7];
			relationshipsOut = (object)paramsArray[8];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="unknownDropGoo">object UnknownDropGoo</param>
		/// <param name="recordSetDefName">string RecordSetDefName</param>
		/// <param name="location">NetOffice.OWC10Api.Enums.DscDropLocationEnum Location</param>
		/// <param name="type">NetOffice.OWC10Api.Enums.DscDropTypeEnum Type</param>
		/// <param name="dropRowsource">string DropRowsource</param>
		/// <param name="rowsourcesOut">object RowsourcesOut</param>
		/// <param name="relationshipsOut">object RelationshipsOut</param>
		/// <param name="numberOfDrops">Int32 NumberOfDrops</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ValidateDrop(object unknownDropGoo, string recordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum location, NetOffice.OWC10Api.Enums.DscDropTypeEnum type, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut, out Int32 numberOfDrops)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,false,false,true,true,true,true);
			dropRowsource = string.Empty;
			rowsourcesOut = null;
			relationshipsOut = null;
			numberOfDrops = 0;
			object[] paramsArray = Invoker.ValidateParamsArray(unknownDropGoo, recordSetDefName, location, type, dropRowsource, rowsourcesOut, relationshipsOut, numberOfDrops);
			Invoker.Method(this, "ValidateDrop", paramsArray, modifiers);
			dropRowsource = (string)paramsArray[4];
			rowsourcesOut = (object)paramsArray[5];
			relationshipsOut = (object)paramsArray[6];
			numberOfDrops = (Int32)paramsArray[7];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="hyperlink">object Hyperlink</param>
		/// <param name="part">NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum Part</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public string HyperlinkPart(object hyperlink, NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(hyperlink, part);
			object returnItem = Invoker.MethodReturn(this, "HyperlinkPart", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SchemaRefresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "SchemaRefresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="oldID">string OldID</param>
		/// <param name="newID">string NewID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void UpdateElementID(string oldID, string newID)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(oldID, newID);
			Invoker.Method(this, "UpdateElementID", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Reset()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Reset", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lIndex">Int32 lIndex</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public string getDataMemberName(Int32 lIndex)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lIndex);
			object returnItem = Invoker.MethodReturn(this, "getDataMemberName", paramsArray);
			return NetRuntimeSystem.Convert.ToString(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 getDataMemberCount()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "getDataMemberCount", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="sectionElement">object SectionElement</param>
		/// <param name="recordSource">string RecordSource</param>
		/// <param name="sectionType">NetOffice.OWC10Api.Enums.SectTypeEnum SectionType</param>
		/// <param name="groupLevel">NetOffice.OWC10Api.GroupLevel GroupLevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void GetSectionInfo(object sectionElement, out string recordSource, out NetOffice.OWC10Api.Enums.SectTypeEnum sectionType, out NetOffice.OWC10Api.GroupLevel groupLevel)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true,true,true);
			recordSource = string.Empty;
			sectionType = 0;
			groupLevel = null;
			object[] paramsArray = Invoker.ValidateParamsArray(sectionElement, recordSource, sectionType, groupLevel);
			Invoker.Method(this, "GetSectionInfo", paramsArray, modifiers);
			recordSource = (string)paramsArray[1];
			sectionType = (NetOffice.OWC10Api.Enums.SectTypeEnum)paramsArray[2];
			groupLevel = (NetOffice.OWC10Api.GroupLevel)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordSource">string RecordSource</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DeleteRecordSourceIfUnused(string recordSource)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordSource);
			Invoker.Method(this, "DeleteRecordSourceIfUnused", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordSource">string RecordSource</param>
		/// <param name="pageField">string PageField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DeletePageFieldIfUnused(string recordSource, string pageField)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordSource, pageField);
			Invoker.Method(this, "DeletePageFieldIfUnused", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="bstrRecordset">string bstrRecordset</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ResetRecordset(string bstrRecordset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(bstrRecordset);
			Invoker.Method(this, "ResetRecordset", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="exportType">NetOffice.OWC10Api.Enums.ExportableConnectStringEnum ExportType</param>
		/// <param name="connectString">string ConnectString</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void GetExportableConnectString(NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType, out string connectString)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,true);
			connectString = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(exportType, connectString);
			Invoker.Method(this, "GetExportableConnectString", paramsArray, modifiers);
			connectString = (string)paramsArray[1];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="eEncoding">optional NetOffice.OWC10Api.Enums.DscEncodingEnum eEncoding = 0</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void ExportXML(object eEncoding)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(eEncoding);
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void ExportXML()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "ExportXML", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="recordsetName">string RecordsetName</param>
		/// <param name="recordset">NetOffice.ADODBApi.Recordset Recordset</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetRootRecordset(string recordsetName, NetOffice.ADODBApi.Recordset recordset)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(recordsetName, recordset);
			Invoker.Method(this, "SetRootRecordset", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="onlineServer">string OnlineServer</param>
		/// <param name="onlineDatabase">string OnlineDatabase</param>
		/// <param name="offlineServer">string OfflineServer</param>
		/// <param name="offlineDatabase">string OfflineDatabase</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public void GetOfflineDisplayInfo(out string onlineServer, out string onlineDatabase, out string offlineServer, out string offlineDatabase)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true,true,true,true);
			onlineServer = string.Empty;
			onlineDatabase = string.Empty;
			offlineServer = string.Empty;
			offlineDatabase = string.Empty;
			object[] paramsArray = Invoker.ValidateParamsArray(onlineServer, onlineDatabase, offlineServer, offlineDatabase);
			Invoker.Method(this, "GetOfflineDisplayInfo", paramsArray, modifiers);
			onlineServer = (string)paramsArray[0];
			onlineDatabase = (string)paramsArray[1];
			offlineServer = (string)paramsArray[2];
			offlineDatabase = (string)paramsArray[3];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="refreshType">optional NetOffice.OWC10Api.Enums.RefreshType RefreshType = 1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Refresh(object refreshType)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(refreshType);
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public void Refresh()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "Refresh", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="fChild">Int32 fChild</param>
		/// <param name="ppGrouplevel">NetOffice.OWC10Api.GroupLevel ppGrouplevel</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void FindRelatedGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, Int32 fChild, out NetOffice.OWC10Api.GroupLevel ppGrouplevel)
		{
			ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false,false,true);
			ppGrouplevel = null;
			object[] paramsArray = Invoker.ValidateParamsArray(pGroupLevel, fChild, ppGrouplevel);
			Invoker.Method(this, "FindRelatedGroupLevel", paramsArray, modifiers);
			ppGrouplevel = (NetOffice.OWC10Api.GroupLevel)paramsArray[2];
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="notification">NetOffice.OWC10Api.Enums.NotificationType Notification</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void DllNotification(NetOffice.OWC10Api.Enums.NotificationType notification)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(notification);
			Invoker.Method(this, "DllNotification", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="suspend">bool Suspend</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SuspendUndo(bool suspend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(suspend);
			Invoker.Method(this, "SuspendUndo", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void UpdateFocus()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "UpdateFocus", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="connectionString">string ConnectionString</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public bool IsValidDAPProvider(string connectionString)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(connectionString);
			object returnItem = Invoker.MethodReturn(this, "IsValidDAPProvider", paramsArray);
			return NetRuntimeSystem.Convert.ToBoolean(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		/// <param name="fullPrecision">optional object FullPrecision</param>
		/// <param name="triangulationPrecision">optional object TriangulationPrecision</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="number">Double Number</param>
		/// <param name="sourceCurrency">string SourceCurrency</param>
		/// <param name="targetCurrency">string TargetCurrency</param>
		/// <param name="fullPrecision">optional object FullPrecision</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(number, sourceCurrency, targetCurrency, fullPrecision);
			object returnItem = Invoker.MethodReturn(this, "EuroConvert", paramsArray);
			return NetRuntimeSystem.Convert.ToDouble(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public String[] GetDAPProviders()
		{
			object[] paramsArray = null;
			object returnItem = (object)Invoker.MethodReturn(this, "GetDAPProviders", paramsArray);
			return (String[])returnItem;
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="synchronizing">bool Synchronizing</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetSynchronizing(bool synchronizing)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(synchronizing);
			Invoker.Method(this, "SetSynchronizing", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="displayError">bool DisplayError</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetDisplayError(bool displayError)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(displayError);
			Invoker.Method(this, "SetDisplayError", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="suspend">bool Suspend</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SuspendXMLReExecute(bool suspend)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(suspend);
			Invoker.Method(this, "SuspendXMLReExecute", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="firePropChange">bool FirePropChange</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void SetFirePropChange(bool firePropChange)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(firePropChange);
			Invoker.Method(this, "SetFirePropChange", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="value">object Value</param>
		/// <param name="valueIfNull">optional object ValueIfNull</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public object Nz(object value, object valueIfNull)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value, valueIfNull);
			object returnItem = Invoker.MethodReturn(this, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="value">object Value</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("OWC10", 1)]
		public object Nz(object value)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(value);
			object returnItem = Invoker.MethodReturn(this, "Nz", paramsArray);
			if((null != returnItem) && (returnItem is MarshalByRefObject))
			{
				ICOMObject newObject = Factory.CreateObjectFromComProxy(this, returnItem);
				return newObject;
			}
			else
			{
				return  returnItem;
			}
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public void RefreshJetCache()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "RefreshJetCache", paramsArray);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersionAttribute("OWC10", 1)]
		public void AutoRefreshOfflineSource()
		{
			object[] paramsArray = null;
			Invoker.Method(this, "AutoRefreshOfflineSource", paramsArray);
		}

		#endregion
		#pragma warning restore
	}
}