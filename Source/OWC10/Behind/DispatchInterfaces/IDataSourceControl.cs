using System.Reflection;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api.Behind
{
    /// <summary>
    /// IDataSourceControl
    /// </summary>
    [SyntaxBypass]
    public class IDataSourceControl_ : COMObject, NetOffice.OWC10Api.IDataSourceControl_
    {
        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IDataSourceControl_() : base()
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
        public virtual NetOffice.OWC10Api.Enums.ProviderType get_ProviderType(object dataMember)
        {
            return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType", dataMember);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Alias for get_ProviderType
        /// </summary>
        /// <param name="dataMember">optional object dataMember</param>
        [SupportByVersion("OWC10", 1), Redirect("get_ProviderType")]
        public virtual NetOffice.OWC10Api.Enums.ProviderType ProviderType(object dataMember)
        {
            return get_ProviderType(dataMember);        }

        #endregion
    }

    /// <summary>
    /// DispatchInterface IDataSourceControl 
    /// SupportByVersion OWC10, 1
    /// </summary>
    [SupportByVersion("OWC10", 1)]
    [EntityType(EntityType.IsDispatchInterface), BaseType]
    public class IDataSourceControl : IDataSourceControl_, NetOffice.OWC10Api.IDataSourceControl
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
                    _contractType = typeof(NetOffice.OWC10Api.IDataSourceControl);
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
                    _type = typeof(IDataSourceControl);
                return _type;
            }
        }

        #endregion

        #region Ctor

        /// <summary>
        /// Stub Ctor, not intended to use
        /// </summary>
        public IDataSourceControl() : base()
        {
        }

        #endregion

        #region Properties

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string ConnectionString
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectionString");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectionString", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string CurrentDirectory
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "CurrentDirectory");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool UseRemoteProvider
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseRemoteProvider");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseRemoteProvider", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.ADODBApi.Connection Connection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Connection>(this, "Connection", typeof(NetOffice.ADODBApi.Connection));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool DataEntry
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DataEntry");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DataEntry", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 MaxRecords
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MaxRecords");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "MaxRecords", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.ADODBApi.Recordset DefaultRecordset
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ADODBApi.Recordset>(this, "DefaultRecordset", typeof(NetOffice.ADODBApi.Recordset));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.SchemaRowsources SchemaRowsources
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRowsources>(this, "SchemaRowsources", typeof(NetOffice.OWC10Api.SchemaRowsources));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.SchemaRelationships SchemaRelationships
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaRelationships>(this, "SchemaRelationships", typeof(NetOffice.OWC10Api.SchemaRelationships));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api.PageRowsources PageRowsources
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.PageRowsources>(this, "PageRowsources", typeof(NetOffice.OWC10Api.PageRowsources));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.RecordsetDefs RecordsetDefs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDefs>(this, "RecordsetDefs", typeof(NetOffice.OWC10Api.RecordsetDefs));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api.RecordsetDefs RootRecordsetDefs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.RecordsetDefs>(this, "RootRecordsetDefs", typeof(NetOffice.OWC10Api.RecordsetDefs));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object PivotDefs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "PivotDefs");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string DefaultRecordsetName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultRecordsetName");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DefaultRecordsetName", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string XMLData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLData", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.GroupLevels GroupLevels
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.GroupLevels>(this, "GroupLevels", typeof(NetOffice.OWC10Api.GroupLevels));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("OWC10", 1), ProxyResult]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual object Constants
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Constants");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.ElementExtensions ElementExtensions
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.ElementExtensions>(this, "ElementExtensions", typeof(NetOffice.OWC10Api.ElementExtensions));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool IsNew
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsNew");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsNew", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum RecordsetType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscRecordsetTypeEnum>(this, "RecordsetType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "RecordsetType", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.AllPageFields AllPageFields
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.AllPageFields>(this, "AllPageFields", typeof(NetOffice.OWC10Api.AllPageFields));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Section CurrentSection
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.Section>(this, "CurrentSection", typeof(NetOffice.OWC10Api.Section));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api.Enums.ProviderType ProviderType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.ProviderType>(this, "ProviderType");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.AllGroupingDefs AllGroupingDefs
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.AllGroupingDefs>(this, "AllGroupingDefs", typeof(NetOffice.OWC10Api.AllGroupingDefs));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool DisplayAlerts
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "DisplayAlerts");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "DisplayAlerts", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.DataPages DataPages
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.DataPages>(this, "DataPages", typeof(NetOffice.OWC10Api.DataPages));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 GridX
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridX");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridX", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 GridY
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "GridY");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "GridY", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual Int32 LoadError
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "LoadError");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Enums.DefaultControlTypeEnum DefaultControlType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DefaultControlTypeEnum>(this, "DefaultControlType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "DefaultControlType", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool IsDirty
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDirty");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsDirty", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool Busy
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Busy");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string Version
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "Version");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 MajorVersion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "MajorVersion");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string MinorVersion
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "MinorVersion");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string BuildNumber
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "BuildNumber");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string RevisionNumber
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "RevisionNumber");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual bool IsDataModelDirty
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "IsDataModelDirty");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "IsDataModelDirty", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Enums.DscOfflineTypeEnum OfflineType
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscOfflineTypeEnum>(this, "OfflineType");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "OfflineType", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string OfflinePublication
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OfflinePublication");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OfflinePublication", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool Offline
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "Offline");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string OfflineSource
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "OfflineSource");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "OfflineSource", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Enums.DscXMLLocationEnum XMLLocation
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.OWC10Api.Enums.DscXMLLocationEnum>(this, "XMLLocation");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteEnumPropertySet(this, "XMLLocation", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual bool UseXMLData
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteBoolPropertyGet(this, "UseXMLData");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "UseXMLData", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string XMLDataTarget
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "XMLDataTarget");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "XMLDataTarget", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get/Set
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual string ConnectionFile
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectionFile");
            }
            set
            {
                InvokerService.InvokeInternal.ExecuteValuePropertySet(this, "ConnectionFile", value);
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string DefaultRecordsetDefName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DefaultRecordsetDefName");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string ConnectionStringFullPath
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "ConnectionStringFullPath");
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual NetOffice.OWC10Api.SchemaDiagrams SchemaDiagrams
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.OWC10Api.SchemaDiagrams>(this, "SchemaDiagrams", typeof(NetOffice.OWC10Api.SchemaDiagrams));
            }
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// Get
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public virtual string DBNSOwnerName
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteStringPropertyGet(this, "DBNSOwnerName");
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
        public virtual NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption, object fetchType)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", typeof(NetOffice.ADODBApi.Recordset), recordsetName, executeOption, fetchType);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.ADODBApi.Recordset Execute(string recordsetName)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", typeof(NetOffice.ADODBApi.Recordset), recordsetName);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        /// <param name="executeOption">optional NetOffice.ADODBApi.Enums.ExecuteOptionEnum ExecuteOption = -1</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.ADODBApi.Recordset Execute(string recordsetName, object executeOption)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ADODBApi.Recordset>(this, "Execute", typeof(NetOffice.ADODBApi.Recordset), recordsetName, executeOption);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="dataAssistant">object dataAssistant</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SetDataAssistant(object dataAssistant)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDataAssistant", dataAssistant);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="advise">object advise</param>
        /// <param name="sinkName">string sinkName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void DesignAdvise(object advise, string sinkName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DesignAdvise", advise, sinkName);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="sinkName">string sinkName</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void DesignUnAdvise(string sinkName)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DesignUnAdvise", sinkName);
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
        public virtual void ProcessDrop(object pUnknownDropGoo, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ProcessDrop", new object[] { pUnknownDropGoo, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship });
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
        public virtual void ScriptDrop(object rowsources, object relationships, object fields, string bstrRecordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, string pageRowsource, string schemaRelationship)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ScriptDrop", new object[] { rowsources, relationships, fields, bstrRecordSetDefName, dl, dt, pageRowsource, schemaRelationship });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="element">object element</param>
        [SupportByVersion("OWC10", 1)]
        public virtual NetOffice.OWC10Api.Section GetContainingSection(object element)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.OWC10Api.Section>(this, "GetContainingSection", typeof(NetOffice.OWC10Api.Section), element);
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
        public virtual void ScriptValidate(object rowsources, object relationships, object fields, string recordsetDef, NetOffice.OWC10Api.Enums.DscDropLocationEnum dl, NetOffice.OWC10Api.Enums.DscDropTypeEnum dt, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, false, false, false, true, true, true);
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
        public virtual void ValidateDrop(object unknownDropGoo, string recordSetDefName, NetOffice.OWC10Api.Enums.DscDropLocationEnum location, NetOffice.OWC10Api.Enums.DscDropTypeEnum type, out string dropRowsource, out object rowsourcesOut, out object relationshipsOut, out Int32 numberOfDrops)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, false, false, true, true, true, true);
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
        public virtual string HyperlinkPart(object hyperlink, NetOffice.OWC10Api.Enums.DscHyperlinkPartEnum part)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "HyperlinkPart", hyperlink, part);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SchemaRefresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SchemaRefresh");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="oldID">string oldID</param>
        /// <param name="newID">string newID</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void UpdateElementID(string oldID, string newID)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateElementID", oldID, newID);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void Reset()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Reset");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="lIndex">Int32 lIndex</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual string getDataMemberName(Int32 lIndex)
        {
            return InvokerService.InvokeInternal.ExecuteStringMethodGet(this, "getDataMemberName", lIndex);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual Int32 getDataMemberCount()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "getDataMemberCount");
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
        public virtual void GetSectionInfo(object sectionElement, out string recordSource, out NetOffice.OWC10Api.Enums.SectTypeEnum sectionType, out NetOffice.OWC10Api.GroupLevel groupLevel)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true, true, true);
            recordSource = string.Empty;
            sectionType = 0;
            groupLevel = null;
            object[] paramsArray = Invoker.ValidateParamsArray(sectionElement, recordSource, sectionType, groupLevel);
            Invoker.Method(this, "GetSectionInfo", paramsArray, modifiers);
            recordSource = paramsArray[1] as string;
            sectionType = (NetOffice.OWC10Api.Enums.SectTypeEnum)paramsArray[2];
            if (paramsArray[3] is MarshalByRefObject)
                groupLevel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.GroupLevel>(this, paramsArray[3], typeof(NetOffice.OWC10Api.GroupLevel));
            else
                groupLevel = null;           
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordSource">string recordSource</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void DeleteRecordSourceIfUnused(string recordSource)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeleteRecordSourceIfUnused", recordSource);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordSource">string recordSource</param>
        /// <param name="pageField">string pageField</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void DeletePageFieldIfUnused(string recordSource, string pageField)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DeletePageFieldIfUnused", recordSource, pageField);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="bstrRecordset">string bstrRecordset</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void ResetRecordset(string bstrRecordset)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ResetRecordset", bstrRecordset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="exportType">NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType</param>
        /// <param name="connectString">string connectString</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void GetExportableConnectString(NetOffice.OWC10Api.Enums.ExportableConnectStringEnum exportType, out string connectString)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, true);
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
        public virtual void ExportXML(object eEncoding)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML", eEncoding);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void ExportXML()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "ExportXML");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="recordsetName">string recordsetName</param>
        /// <param name="recordset">NetOffice.ADODBApi.Recordset recordset</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void SetRootRecordset(string recordsetName, NetOffice.ADODBApi.Recordset recordset)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetRootRecordset", recordsetName, recordset);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="onlineServer">string onlineServer</param>
        /// <param name="onlineDatabase">string onlineDatabase</param>
        /// <param name="offlineServer">string offlineServer</param>
        /// <param name="offlineDatabase">string offlineDatabase</param>
        [SupportByVersion("OWC10", 1)]
        public virtual void GetOfflineDisplayInfo(out string onlineServer, out string onlineDatabase, out string offlineServer, out string offlineDatabase)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(true, true, true, true);
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
        public virtual void Refresh(object refreshType)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh", refreshType);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual void Refresh()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "Refresh");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
        /// <param name="fChild">Int32 fChild</param>
        /// <param name="ppGrouplevel">NetOffice.OWC10Api.GroupLevel ppGrouplevel</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void FindRelatedGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, Int32 fChild, out NetOffice.OWC10Api.GroupLevel ppGrouplevel)
        {
            ParameterModifier[] modifiers = Invoker.CreateParamModifiers(false, false, true);
            ppGrouplevel = null;
            object[] paramsArray = Invoker.ValidateParamsArray(pGroupLevel, fChild, ppGrouplevel);
            Invoker.Method(this, "FindRelatedGroupLevel", paramsArray, modifiers);
            if (paramsArray[2] is MarshalByRefObject)
                ppGrouplevel = Factory.CreateKnownObjectFromComProxy<NetOffice.OWC10Api.GroupLevel>(this, paramsArray[2], typeof(NetOffice.OWC10Api.GroupLevel));
            else
                ppGrouplevel = null;

        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="notification">NetOffice.OWC10Api.Enums.NotificationType notification</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void DllNotification(NetOffice.OWC10Api.Enums.NotificationType notification)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "DllNotification", notification);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="suspend">bool suspend</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SuspendUndo(bool suspend)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SuspendUndo", suspend);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void UpdateFocus()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "UpdateFocus");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="connectionString">string connectionString</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual bool IsValidDAPProvider(string connectionString)
        {
            return InvokerService.InvokeInternal.ExecuteBoolMethodGet(this, "IsValidDAPProvider", connectionString);
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
        public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision, object triangulationPrecision)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", new object[] { number, sourceCurrency, targetCurrency, fullPrecision, triangulationPrecision });
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="number">Double number</param>
        /// <param name="sourceCurrency">string sourceCurrency</param>
        /// <param name="targetCurrency">string targetCurrency</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency);
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
        public virtual Double EuroConvert(Double number, string sourceCurrency, string targetCurrency, object fullPrecision)
        {
            return InvokerService.InvokeInternal.ExecuteDoubleMethodGet(this, "EuroConvert", number, sourceCurrency, targetCurrency, fullPrecision);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual String[] GetDAPProviders()
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
        public virtual void SetSynchronizing(bool synchronizing)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetSynchronizing", synchronizing);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="displayError">bool displayError</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SetDisplayError(bool displayError)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetDisplayError", displayError);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="suspend">bool suspend</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SuspendXMLReExecute(bool suspend)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SuspendXMLReExecute", suspend);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="firePropChange">bool firePropChange</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void SetFirePropChange(bool firePropChange)
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "SetFirePropChange", firePropChange);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="value">object value</param>
        /// <param name="valueIfNull">optional object valueIfNull</param>
        [SupportByVersion("OWC10", 1)]
        public virtual object Nz(object value, object valueIfNull)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Nz", value, valueIfNull);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        /// <param name="value">object value</param>
        [CustomMethod]
        [SupportByVersion("OWC10", 1)]
        public virtual object Nz(object value)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Nz", value);
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [SupportByVersion("OWC10", 1)]
        public virtual void RefreshJetCache()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "RefreshJetCache");
        }

        /// <summary>
        /// SupportByVersion OWC10 1
        /// </summary>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        [SupportByVersion("OWC10", 1)]
        public virtual void AutoRefreshOfflineSource()
        {
            InvokerService.InvokeInternal.ExecuteMethod(this, "AutoRefreshOfflineSource");
        }

        #endregion

        #pragma warning restore
    }
}
