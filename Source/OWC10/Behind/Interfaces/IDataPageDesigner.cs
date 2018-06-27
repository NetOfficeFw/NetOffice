using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.OWC10Api;

namespace NetOffice.OWC10Api.Behind
{
	/// <summary>
	/// Interface IDataPageDesigner 
	/// SupportByVersion OWC10, 1
	/// </summary>
	[SupportByVersion("OWC10", 1)]
	[EntityType(EntityType.IsInterface)]
 	public class IDataPageDesigner : COMObject, NetOffice.OWC10Api.IDataPageDesigner
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
                    _contractType = typeof(NetOffice.OWC10Api.IDataPageDesigner);
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
                    _type = typeof(IDataPageDesigner);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IDataPageDesigner() : base()
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
		public virtual Int32 ConnectDataComponents(NetOffice.OWC10Api.IDataSourceControl pDataSourceControl)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "ConnectDataComponents", pDataSourceControl);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 CreateSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "CreateSection", sectType, wzRecordsetName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum sectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		/// <param name="fInGroupingDefDelete">Int32 fInGroupingDefDelete</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 DeleteSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName, Int32 fInGroupingDefDelete)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DeleteSection", sectType, wzRecordsetName, fInGroupingDefDelete);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 OnGroupLevelAdded(NetOffice.OWC10Api.GroupLevel pGroupLevel)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnGroupLevelAdded", pGroupLevel);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 OnGroupLevelDeleted()
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnGroupLevelDeleted");
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="wzRecordsetNameOld">string wzRecordsetNameOld</param>
		/// <param name="wzRecordsetNameNew">string wzRecordsetNameNew</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 RebindGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, string wzRecordsetNameOld, string wzRecordsetNameNew)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RebindGroupLevel", pGroupLevel, wzRecordsetNameOld, wzRecordsetNameNew);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedConnection">object ppUnknownSharedConnection</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetSharedConnectionObject(string wzConnectionString, object ppUnknownSharedConnection)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetSharedConnectionObject", wzConnectionString, ppUnknownSharedConnection);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="lMarker">Int32 lMarker</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 TWPerformanceMarker(Int32 lMarker)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "TWPerformanceMarker", lMarker);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="pfSecure">Int32 pfSecure</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 IsDatabaseSecure(string wzConnectionString, Int32 pfSecure)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsDatabaseSecure", wzConnectionString, pfSecure);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="dispidChanged">Int32 dispidChanged</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 OnPropChanged(Int32 dispidChanged)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "OnPropChanged", dispidChanged);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedDBNS">object ppUnknownSharedDBNS</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetSharedDBNS(string wzConnectionString, object ppUnknownSharedDBNS)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetSharedDBNS", wzConnectionString, ppUnknownSharedDBNS);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrFileName">string ppbstrFileName</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetDatapagePath(string ppbstrFileName)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetDatapagePath", ppbstrFileName);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pfDesignMode">Int32 pfDesignMode</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 IsDesignMode(Int32 pfDesignMode)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "IsDesignMode", pfDesignMode);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pRequestingDSC">NetOffice.OWC10Api.IDataSourceControl pRequestingDSC</param>
		/// <param name="vfForceRefresh">bool vfForceRefresh</param>
		/// <param name="rt">NetOffice.OWC10Api.Enums.RefreshType rt</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 RefreshDataTools(NetOffice.OWC10Api.IDataSourceControl pRequestingDSC, bool vfForceRefresh, NetOffice.OWC10Api.Enums.RefreshType rt)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "RefreshDataTools", pRequestingDSC, vfForceRefresh, rt);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="ppbstrInstId">string ppbstrInstId</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetFieldListInstanceId(string ppbstrInstId)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetFieldListInstanceId", ppbstrInstId);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// </summary>
		/// <param name="pioum">NetOffice.OWC10Api.IOleUndoManager pioum</param>
		[SupportByVersion("OWC10", 1)]
		public virtual Int32 GetUndoManager(NetOffice.OWC10Api.IOleUndoManager pioum)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "GetUndoManager", pioum);
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
		public virtual Int32 DoRelWiz(NetOffice.OWC10Api.IDataSourceControl pDSC, string bstrRecordSetDef, string bstrDropRowsource, object varRowsources, object varRelationships, NetOffice.OWC10Api.PageRowsource ppprs, NetOffice.OWC10Api.SchemaRelationship ppsr)
		{
			return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "DoRelWiz", new object[]{ pDSC, bstrRecordSetDef, bstrDropRowsource, varRowsources, varRelationships, ppprs, ppsr });
		}

		#endregion

		#pragma warning restore
	}
}

