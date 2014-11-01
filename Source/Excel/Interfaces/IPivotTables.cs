using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using System.Collections;
using NetOffice;
namespace NetOffice.ExcelApi
{
	///<summary>
	/// Interface IPivotTables 
	/// SupportByVersion Excel, 9,10,11,12,14,15
	///</summary>
	[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
	[EntityTypeAttribute(EntityType.IsInterface)]
	public class IPivotTables : COMObject ,IEnumerable<NetOffice.ExcelApi.PivotTable>
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
                    _type = typeof(IPivotTables);
                    
                return _type;
            }
        }
        
        #endregion
        
		#region Construction

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IPivotTables(Core factory, COMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables(COMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables(Core factory, COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables(COMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables(COMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables() : base()
		{
		}
		
		/// <param name="progId">registered ProgID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IPivotTables(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Application", paramsArray);
				NetOffice.ExcelApi.Application newObject = Factory.CreateKnownObjectFromComProxy(this,returnItem,NetOffice.ExcelApi.Application.LateBindingApiWrapperType) as NetOffice.ExcelApi.Application;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Creator", paramsArray);
				int intReturnItem = NetRuntimeSystem.Convert.ToInt32(returnItem);
				return (NetOffice.ExcelApi.Enums.XlCreator)intReturnItem;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public object Parent
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Parent", paramsArray);
				COMObject newObject = Factory.CreateObjectFromComProxy(this,returnItem);
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// Get
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public Int32 Count
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "Count", paramsArray);
				return NetRuntimeSystem.Convert.ToInt32(returnItem);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// 
		/// </summary>
		/// <param name="index">object Index</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item")]
		public NetOffice.ExcelApi.PivotTable this[object index]
		{
			get
			{
				object[] paramsArray = Invoker.ValidateParamsArray(index);
				object returnItem = Invoker.MethodReturn(this, "Item", paramsArray);
				NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
				return newObject;
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// 
		/// </summary>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache PivotCache</param>
		/// <param name="tableDestination">object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="readData">optional object ReadData</param>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName, object readData)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pivotCache, tableDestination, tableName, readData);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15
		/// 
		/// </summary>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache PivotCache</param>
		/// <param name="tableDestination">object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		/// <param name="readData">optional object ReadData</param>
		/// <param name="defaultVersion">optional object DefaultVersion</param>
		[SupportByVersionAttribute("Excel", 10,11,12,14,15)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName, object readData, object defaultVersion)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pivotCache, tableDestination, tableName, readData, defaultVersion);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// 
		/// </summary>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache PivotCache</param>
		/// <param name="tableDestination">object TableDestination</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pivotCache, tableDestination);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
			return newObject;
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15
		/// 
		/// </summary>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache PivotCache</param>
		/// <param name="tableDestination">object TableDestination</param>
		/// <param name="tableName">optional object TableName</param>
		[CustomMethodAttribute]
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName)
		{
			object[] paramsArray = Invoker.ValidateParamsArray(pivotCache, tableDestination, tableName);
			object returnItem = Invoker.MethodReturn(this, "Add", paramsArray);
			NetOffice.ExcelApi.PivotTable newObject = Factory.CreateKnownObjectFromComProxy(this, returnItem,NetOffice.ExcelApi.PivotTable.LateBindingApiWrapperType) as NetOffice.ExcelApi.PivotTable;
			return newObject;
		}

		#endregion

       #region IEnumerable<NetOffice.ExcelApi.PivotTable> Member
        
        /// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
       public IEnumerator<NetOffice.ExcelApi.PivotTable> GetEnumerator()  
       {
           NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
           foreach (NetOffice.ExcelApi.PivotTable item in innerEnumerator)
               yield return item;
       }

       #endregion
          
		#region IEnumerable Members
       
		/// <summary>
		/// SupportByVersionAttribute Excel, 9,10,11,12,14,15
		/// </summary>
		[SupportByVersionAttribute("Excel", 9,10,11,12,14,15)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this);
		}

		#endregion
		#pragma warning restore
	}
}