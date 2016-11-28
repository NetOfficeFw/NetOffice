using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OWC10Api
{
	///<summary>
	/// Interface IDataPageDesigner 
	/// SupportByVersion OWC10, 1
	///</summary>
	[SupportByVersionAttribute("OWC10", 1)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IDataPageDesigner : COMObject
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
                    _type = typeof(IDataPageDesigner);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

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
		
		/// <param name="progId">registered ProgID</param>
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
		/// 
		/// </summary>
		/// <param name="pDataSourceControl">NetOffice.OWC10Api.IDataSourceControl pDataSourceControl</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 ConnectDataComponents(NetOffice.OWC10Api.IDataSourceControl pDataSourceControl)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pDataSourceControl);
			object returnItem = Invoker.MethodReturn(this, "ConnectDataComponents", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum SectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 CreateSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectType, wzRecordsetName);
			object returnItem = Invoker.MethodReturn(this, "CreateSection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="sectType">NetOffice.OWC10Api.Enums.SectTypeEnum SectType</param>
		/// <param name="wzRecordsetName">string wzRecordsetName</param>
		/// <param name="fInGroupingDefDelete">Int32 fInGroupingDefDelete</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 DeleteSection(NetOffice.OWC10Api.Enums.SectTypeEnum sectType, string wzRecordsetName, Int32 fInGroupingDefDelete)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(sectType, wzRecordsetName, fInGroupingDefDelete);
			object returnItem = Invoker.MethodReturn(this, "DeleteSection", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 OnGroupLevelAdded(NetOffice.OWC10Api.GroupLevel pGroupLevel)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pGroupLevel);
			object returnItem = Invoker.MethodReturn(this, "OnGroupLevelAdded", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 OnGroupLevelDeleted()
		{
			object[] paramsArray = null;
			object returnItem = Invoker.MethodReturn(this, "OnGroupLevelDeleted", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pGroupLevel">NetOffice.OWC10Api.GroupLevel pGroupLevel</param>
		/// <param name="wzRecordsetNameOld">string wzRecordsetNameOld</param>
		/// <param name="wzRecordsetNameNew">string wzRecordsetNameNew</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 RebindGroupLevel(NetOffice.OWC10Api.GroupLevel pGroupLevel, string wzRecordsetNameOld, string wzRecordsetNameNew)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pGroupLevel, wzRecordsetNameOld, wzRecordsetNameNew);
			object returnItem = Invoker.MethodReturn(this, "RebindGroupLevel", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedConnection">object ppUnknownSharedConnection</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetSharedConnectionObject(string wzConnectionString, object ppUnknownSharedConnection)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wzConnectionString, ppUnknownSharedConnection);
			object returnItem = Invoker.MethodReturn(this, "GetSharedConnectionObject", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="lMarker">Int32 lMarker</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 TWPerformanceMarker(Int32 lMarker)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(lMarker);
			object returnItem = Invoker.MethodReturn(this, "TWPerformanceMarker", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="pfSecure">Int32 pfSecure</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 IsDatabaseSecure(string wzConnectionString, Int32 pfSecure)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wzConnectionString, pfSecure);
			object returnItem = Invoker.MethodReturn(this, "IsDatabaseSecure", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="dispidChanged">Int32 dispidChanged</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 OnPropChanged(Int32 dispidChanged)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(dispidChanged);
			object returnItem = Invoker.MethodReturn(this, "OnPropChanged", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="wzConnectionString">string wzConnectionString</param>
		/// <param name="ppUnknownSharedDBNS">object ppUnknownSharedDBNS</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetSharedDBNS(string wzConnectionString, object ppUnknownSharedDBNS)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(wzConnectionString, ppUnknownSharedDBNS);
			object returnItem = Invoker.MethodReturn(this, "GetSharedDBNS", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="ppbstrFileName">string ppbstrFileName</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetDatapagePath(string ppbstrFileName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(ppbstrFileName);
			object returnItem = Invoker.MethodReturn(this, "GetDatapagePath", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pfDesignMode">Int32 pfDesignMode</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 IsDesignMode(Int32 pfDesignMode)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pfDesignMode);
			object returnItem = Invoker.MethodReturn(this, "IsDesignMode", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pRequestingDSC">NetOffice.OWC10Api.IDataSourceControl pRequestingDSC</param>
		/// <param name="vfForceRefresh">bool vfForceRefresh</param>
		/// <param name="rt">NetOffice.OWC10Api.Enums.RefreshType rt</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 RefreshDataTools(NetOffice.OWC10Api.IDataSourceControl pRequestingDSC, bool vfForceRefresh, NetOffice.OWC10Api.Enums.RefreshType rt)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pRequestingDSC, vfForceRefresh, rt);
			object returnItem = Invoker.MethodReturn(this, "RefreshDataTools", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="ppbstrInstId">string ppbstrInstId</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetFieldListInstanceId(string ppbstrInstId)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(ppbstrInstId);
			object returnItem = Invoker.MethodReturn(this, "GetFieldListInstanceId", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pioum">NetOffice.OWC10Api.IOleUndoManager pioum</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 GetUndoManager(NetOffice.OWC10Api.IOleUndoManager pioum)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pioum);
			object returnItem = Invoker.MethodReturn(this, "GetUndoManager", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		/// <summary>
		/// SupportByVersion OWC10 1
		/// 
		/// </summary>
		/// <param name="pDSC">NetOffice.OWC10Api.IDataSourceControl pDSC</param>
		/// <param name="bstrRecordSetDef">string bstrRecordSetDef</param>
		/// <param name="bstrDropRowsource">string bstrDropRowsource</param>
		/// <param name="varRowsources">object varRowsources</param>
		/// <param name="varRelationships">object varRelationships</param>
		/// <param name="ppprs">NetOffice.OWC10Api.PageRowsource ppprs</param>
		/// <param name="ppsr">NetOffice.OWC10Api.SchemaRelationship ppsr</param>
		[SupportByVersionAttribute("OWC10", 1)]
		public Int32 DoRelWiz(NetOffice.OWC10Api.IDataSourceControl pDSC, string bstrRecordSetDef, string bstrDropRowsource, object varRowsources, object varRelationships, NetOffice.OWC10Api.PageRowsource ppprs, NetOffice.OWC10Api.SchemaRelationship ppsr)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pDSC, bstrRecordSetDef, bstrDropRowsource, varRowsources, varRelationships, ppprs, ppsr);
			object returnItem = Invoker.MethodReturn(this, "DoRelWiz", paramsArray);
			return NetRuntimeSystem.Convert.ToInt32(returnItem);
		}

		#endregion
		#pragma warning restore
	}
}