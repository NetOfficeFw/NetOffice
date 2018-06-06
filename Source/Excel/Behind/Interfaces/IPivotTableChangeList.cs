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
	/// Interface IPivotTableChangeList 
	/// SupportByVersion Excel, 14,15,16
	/// </summary>	[SupportByVersion("Excel", 14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class IPivotTableChangeList : COMObject, NetOffice.ExcelApi.IPivotTableChangeList
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
                    _type = typeof(IPivotTableChangeList);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IPivotTableChangeList() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.ExcelApi.ValueChange this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.ValueChange>(this, "_Default", typeof(NetOffice.ExcelApi.ValueChange), index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		/// <param name="allocationMethod">optional object allocationMethod</param>
		/// <param name="allocationWeightExpression">optional object allocationWeightExpression</param>
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue, object allocationMethod, object allocationWeightExpression)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ValueChange>(this, "Add", typeof(NetOffice.ExcelApi.ValueChange), new object[]{ tuple, value, allocationValue, allocationMethod, allocationWeightExpression });
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.ValueChange Add(string tuple, Double value)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ValueChange>(this, "Add", typeof(NetOffice.ExcelApi.ValueChange), tuple, value);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ValueChange>(this, "Add", typeof(NetOffice.ExcelApi.ValueChange), tuple, value, allocationValue);
		}

		/// <summary>
		/// SupportByVersion Excel 14, 15, 16
		/// </summary>
		/// <param name="tuple">string tuple</param>
		/// <param name="value">Double value</param>
		/// <param name="allocationValue">optional object allocationValue</param>
		/// <param name="allocationMethod">optional object allocationMethod</param>
		[CustomMethod]
		[SupportByVersion("Excel", 14,15,16)]
		public virtual NetOffice.ExcelApi.ValueChange Add(string tuple, Double value, object allocationValue, object allocationMethod)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.ValueChange>(this, "Add", typeof(NetOffice.ExcelApi.ValueChange), tuple, value, allocationValue, allocationMethod);
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.ValueChange>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.ValueChange>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.ValueChange>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.ValueChange>

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.ValueChange> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.ValueChange item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 14,15,16
        /// </summary>
        [SupportByVersion("Excel", 14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

