using System.Collections;
using System.Collections.Generic;
using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;
using NetOffice.CollectionsGeneric;

namespace NetOffice.ExcelApi
{
	/// <summary>
	/// Interface ISortFields 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class ISortFields : COMObject, IEnumerableProvider<NetOffice.ExcelApi.SortField>
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
                    _type = typeof(ISortFields);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public ISortFields(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public ISortFields(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public ISortFields(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public object Parent
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 Count
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "Count");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.SortField this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.SortField>(this, "_Default", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="key">NetOffice.ExcelApi.Range key</param>
		/// <param name="sortOn">optional object sortOn</param>
		/// <param name="order">optional object order</param>
		/// <param name="customOrder">optional object customOrder</param>
		/// <param name="dataOption">optional object dataOption</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.SortField Add(NetOffice.ExcelApi.Range key, object sortOn, object order, object customOrder, object dataOption)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SortField>(this, "Add", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, new object[]{ key, sortOn, order, customOrder, dataOption });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="key">NetOffice.ExcelApi.Range key</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.SortField Add(NetOffice.ExcelApi.Range key)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SortField>(this, "Add", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, key);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="key">NetOffice.ExcelApi.Range key</param>
		/// <param name="sortOn">optional object sortOn</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.SortField Add(NetOffice.ExcelApi.Range key, object sortOn)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SortField>(this, "Add", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, key, sortOn);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="key">NetOffice.ExcelApi.Range key</param>
		/// <param name="sortOn">optional object sortOn</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.SortField Add(NetOffice.ExcelApi.Range key, object sortOn, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SortField>(this, "Add", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, key, sortOn, order);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="key">NetOffice.ExcelApi.Range key</param>
		/// <param name="sortOn">optional object sortOn</param>
		/// <param name="order">optional object order</param>
		/// <param name="customOrder">optional object customOrder</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.SortField Add(NetOffice.ExcelApi.Range key, object sortOn, object order, object customOrder)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.SortField>(this, "Add", NetOffice.ExcelApi.SortField.LateBindingApiWrapperType, key, sortOn, order, customOrder);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		[SupportByVersion("Excel", 12,14,15,16)]
		public Int32 Clear()
		{
			return Factory.ExecuteInt32MethodGet(this, "Clear");
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.SortField>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.SortField>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.SortField>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.SortField>

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.SortField> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.SortField item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12,14,15,16)]
		IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
		{
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}