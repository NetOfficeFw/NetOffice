using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;
using NetOffice.ExcelApi;

namespace NetOffice.ExcelApi.Behind
{
	/// <summary>
	/// DispatchInterface PivotTables 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff821292.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Method, "Item")]
	public class PivotTables : COMObject, NetOffice.ExcelApi.PivotTables
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
                    _type = typeof(PivotTables);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotTables() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff838196.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff834412.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840741.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff839926.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.PivotTable this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "Item", typeof(NetOffice.ExcelApi.PivotTable), index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196910.aspx </remarks>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache pivotCache</param>
		/// <param name="tableDestination">object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="readData">optional object readData</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName, object readData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "Add", typeof(NetOffice.ExcelApi.PivotTable), pivotCache, tableDestination, tableName, readData);
		}

		/// <summary>
		/// SupportByVersion Excel 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196910.aspx </remarks>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache pivotCache</param>
		/// <param name="tableDestination">object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		/// <param name="readData">optional object readData</param>
		/// <param name="defaultVersion">optional object defaultVersion</param>
		[SupportByVersion("Excel", 10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName, object readData, object defaultVersion)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "Add", typeof(NetOffice.ExcelApi.PivotTable), new object[]{ pivotCache, tableDestination, tableName, readData, defaultVersion });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196910.aspx </remarks>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache pivotCache</param>
		/// <param name="tableDestination">object tableDestination</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "Add", typeof(NetOffice.ExcelApi.PivotTable), pivotCache, tableDestination);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196910.aspx </remarks>
		/// <param name="pivotCache">NetOffice.ExcelApi.PivotCache pivotCache</param>
		/// <param name="tableDestination">object tableDestination</param>
		/// <param name="tableName">optional object tableName</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotTable Add(NetOffice.ExcelApi.PivotCache pivotCache, object tableDestination, object tableName)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotTable>(this, "Add", typeof(NetOffice.ExcelApi.PivotTable), pivotCache, tableDestination, tableName);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.PivotTable>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.PivotTable>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.PivotTable>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.PivotTable>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.PivotTable> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.PivotTable item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsMethod(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

