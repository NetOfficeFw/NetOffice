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
	/// DispatchInterface PivotFilters 
	/// SupportByVersion Excel, 12,14,15,16
	/// </summary>
    /// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841200.aspx </remarks>
	public class PivotFilters : COMObject, NetOffice.ExcelApi.PivotFilters
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
                    _contractType = typeof(NetOffice.ExcelApi.PivotFilters);
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
                    _type = typeof(PivotFilters);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public PivotFilters() : base()
		{

		}

		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff836475.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual NetOffice.ExcelApi.Application Application
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840815.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841209.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16), ProxyResult]
		public virtual object Parent
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		public virtual NetOffice.ExcelApi.PivotFilter this[object index]
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.PivotFilter>(this, "_Default", typeof(NetOffice.ExcelApi.PivotFilter), index);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff837566.aspx </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual Int32 Count
		{
			get
			{
				return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField });
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter, object movingPeriod)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField, wholeDayFilter, movingPeriod });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), type);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		/// <param name="value1">optional object value1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField, value1);
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField, value1, value2);
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order });
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name });
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description });
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
		public virtual NetOffice.ExcelApi.PivotFilter Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField, object wholeDayFilter)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField, wholeDayFilter });
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description, object memberPropertyField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description, memberPropertyField });
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), type);
		}

		/// <summary>
		/// SupportByVersion Excel 15,16
		/// </summary>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlPivotFilterType type</param>
		/// <param name="dataField">optional object dataField</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		[CustomMethod]
		[SupportByVersion("Excel", 15, 16)]
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField);
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField, value1);
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), type, dataField, value1, value2);
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order });
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name });
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
		public virtual NetOffice.ExcelApi.PivotFilter _Add(NetOffice.ExcelApi.Enums.XlPivotFilterType type, object dataField, object value1, object value2, object order, object name, object description)
		{
			return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.PivotFilter>(this, "_Add", typeof(NetOffice.ExcelApi.PivotFilter), new object[]{ type, dataField, value1, value2, order, name, description });
		}

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.PivotFilter>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
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
        public virtual IEnumerator<NetOffice.ExcelApi.PivotFilter> GetEnumerator()
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
		}

		#endregion

		#pragma warning restore
	}
}

