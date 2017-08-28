using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.OWC10Api
{
	/// <summary>
	/// Interface IDataPageDesigner 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IDataPageDesigner : COMObject
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
                    _type = typeof(IDataPageDesigner);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IDataPageDesigner(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IDataPageDesigner(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IDataPageDesigner(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pDataSourceControl">NetOffice.OWC10Api.IDataSourceControl pDataSourceControl</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 ConnectDataComponents(NetOffice.OWC10Api.IDataSourceControl pDataSourceControl)
		{
			return Factory.ExecuteInt32MethodGet(this, "ConnectDataComponents", pDataSourceControl);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 CreateSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName)
		{
			return Factory.ExecuteInt32MethodGet(this, "CreateSection", sectType, wzRecordsetName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		/// <param name="fInGroupingDefDelete">Int32 fInGroupingDefDelete</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 DeleteSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName, Int32 fInGroupingDefDelete)
		{
			return Factory.ExecuteInt32MethodGet(this, "DeleteSection", sectType, wzRecordsetName, fInGroupingDefDelete);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 OnGroupLevelAdded(NetOffice.OWC10Api.GroupLevel pGroupLevel)
		{
			return Factory.ExecuteInt32MethodGet(this, "OnGroupLevelAdded", pGroupLevel);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public Int32 OnGroupLevelDeleted()
		{
			return Factory.ExecuteInt32MethodGet(this, "OnGroupLevelDeleted");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="wzRecordsetNameOld">string wzRecordsetNameOld</param>
		/// <param name="wzRecordsetNameNew">string wzRecordsetNameNew</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 RebindGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, string wzRecordsetNameOld, string wzRecordsetNameNew)
		{
			return Factory.ExecuteInt32MethodGet(this, "RebindGroupLevel", pGroupLevel, wzRecordsetNameOld, wzRecordsetNameNew);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedConnection">object ppUnknownSharedConnection</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetSharedConnectionObject(string wzConnectionString, object ppUnknownSharedConnection)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetSharedConnectionObject", wzConnectionString, ppUnknownSharedConnection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lMarker">Int32 lMarker</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 TWPerformanceMarker(Int32 lMarker)
		{
			return Factory.ExecuteInt32MethodGet(this, "TWPerformanceMarker", lMarker);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="pfSecure">Int32 pfSecure</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 IsDatabaseSecure(string wzConnectionString, Int32 pfSecure)
		{
			return Factory.ExecuteInt32MethodGet(this, "IsDatabaseSecure", wzConnectionString, pfSecure);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dispidChanged">Int32 dispidChanged</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 OnPropChanged(Int32 dispidChanged)
		{
			return Factory.ExecuteInt32MethodGet(this, "OnPropChanged", dispidChanged);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedDBNS">object ppUnknownSharedDBNS</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetSharedDBNS(string wzConnectionString, object ppUnknownSharedDBNS)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetSharedDBNS", wzConnectionString, ppUnknownSharedDBNS);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrFileName">string ppbstrFileName</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetDatapagePath(string ppbstrFileName)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetDatapagePath", ppbstrFileName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfDesignMode">Int32 pfDesignMode</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 IsDesignMode(Int32 pfDesignMode)
		{
			return Factory.ExecuteInt32MethodGet(this, "IsDesignMode", pfDesignMode);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pRequestingDSC">NetOffice.OWC10Api.IDataSourceControl pRequestingDSC</param>
		/// <param name="vfForceRefresh">bool vfForceRefresh</param>
		/// <param name="rt">NetOffice.OWC10Api.Enums.RefreshType rt</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 RefreshDataTools(NetOffice.OWC10Api.IDataSourceControl pRequestingDSC, bool vfForceRefresh, NetOffice.OWC10Api.Enums.RefreshType rt)
		{
			return Factory.ExecuteInt32MethodGet(this, "RefreshDataTools", pRequestingDSC, vfForceRefresh, rt);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrInstId">string ppbstrInstId</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetFieldListInstanceId(string ppbstrInstId)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetFieldListInstanceId", ppbstrInstId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pioum">NetOffice.OWC10Api.IOleUndoManager pioum</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 GetUndoManager(NetOffice.OWC10Api.IOleUndoManager pioum)
		{
			return Factory.ExecuteInt32MethodGet(this, "GetUndoManager", pioum);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pDSC">NetOffice.OWC10Api.IDataSourceControl pDSC</param>
		/// <param name="bstrRecordSetDef">string bstrRecordSetDef</param>
		/// <param name="bstrDropRowsource">string bstrDropRowsource</param>
		/// <param name="varRowsources">object varRowsources</param>
		/// <param name="varRelationships">object varRelationships</param>
		/// <param name="ppprs">NetOffice.OWC10Api.PageRowsource ppprs</param>
		/// <param name="ppsr">NetOffice.OWC10Api.SchemaRelationship ppsr</param>
		[SupportByVersion("OWC10", 1)]
		public Int32 DoRelWiz(NetOffice.OWC10Api.IDataSourceControl pDSC, string bstrRecordSetDef, string bstrDropRowsource, object varRowsources, object varRelationships, NetOffice.OWC10Api.PageRowsource ppprs, NetOffice.OWC10Api.SchemaRelationship ppsr)
		{
			return Factory.ExecuteInt32MethodGet(this, "DoRelWiz", new object[]{ pDSC, bstrRecordSetDef, bstrDropRowsource, varRowsources, varRelationships, ppprs, ppsr });
		}

		#endregion

		#pragma warning restore
	}
}
