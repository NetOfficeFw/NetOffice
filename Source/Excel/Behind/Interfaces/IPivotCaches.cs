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
	/// Interface IPivotCaches 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Method), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class IPivotCaches : COMObject, NetOffice.ExcelApi.IPivotCaches
	{
		#pragma warning disable

		#region Type Information

        /// <summary>
        /// Instance Type        /// </summary>
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
                    _type = typeof(IPivotCaches);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IPivotCaches() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
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
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.PivotCache this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotCache>(this, "_Default", typeof(NetOffice.ExcelApi.PivotCache), index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache Add(NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType, object sourceData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "Add", typeof(NetOffice.ExcelApi.PivotCache), sourceType, sourceData);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache Add(NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "Add", typeof(NetOffice.ExcelApi.PivotCache), sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		/// <param name="version">optional object version</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache Create(NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType, object sourceData, object version)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "Create", typeof(NetOffice.ExcelApi.PivotCache), sourceType, sourceData, version);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache Create(NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "Create", typeof(NetOffice.ExcelApi.PivotCache), sourceType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="sourceType">NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType</param>
		/// <param name="sourceData">optional object sourceData</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotCache Create(NetOffice.ExcelApi.Enums.XlPivotTableSourceType sourceType, object sourceData)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotCache>(this, "Create", typeof(NetOffice.ExcelApi.PivotCache), sourceType, sourceData);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.PivotCache>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.PivotCache>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsMethod(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.PivotCache>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.PivotCache>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.PivotCache> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.PivotCache item in innerEnumerator)
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

