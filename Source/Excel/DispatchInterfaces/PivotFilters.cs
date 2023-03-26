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
	/// DispatchInterface PivotFilters 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotFilters"/> </remarks>
	[SupportByVersion("Excel", 12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class PivotFilters : COMObject, IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>
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
                    _type = typeof(PivotFilters);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public PivotFilters(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public PivotFilters(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public PivotFilters(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotFilters.Application"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotFilters.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotFilters.Parent"/> </remarks>
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
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public NetOffice.ExcelApi.PivotFilter this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFilter>(this, "_Default", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.PivotFilters.Count"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
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
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		/// <param name="wholeDayFilter">optional object wholeDayFilter</param>
		/// <param name="movingPeriod">optional object movingPeriod</param>
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter, object movingPeriod)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField, wholeDayFilter, movingPeriod });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField, value1);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField, value1, value2);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		/// <param name="wholeDayFilter">optional object wholeDayFilter</param>
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField, wholeDayFilter });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		/// <param name="memberPropertyField">optional object memberPropertyField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField, value1);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, type, dataField, value1, value2);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		/// <param name="value2">optional object value2</param>
		/// <param name="order">optional object order</param>
		/// <param name="name">optional object name</param>
		/// <param name="description">optional object description</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", NetOffice.ExcelApi.PivotFilter.LateBindingApiWrapperType, new object[]{ type, dataField, value1, value2, order, name, description });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.PivotFilter>

        /// <summary>
        /// SupportByVersion Excel, 12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public IEnumerator<NetOffice.ExcelApi.PivotFilter> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.PivotFilter item in innerEnumerator)
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