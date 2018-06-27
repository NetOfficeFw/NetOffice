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
    /// Interface IFormatConditions 
    /// SupportByVersion Excel, 9,10,11,12,14,15,16
    /// </summary>
    public class IFormatConditions : COMObject, NetOffice.ExcelApi.IFormatConditions
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
                    _contractType = typeof(NetOffice.ExcelApi.IFormatConditions);
                return _contractType;
            }
        }
        private static Type _contractType;


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
                    _type = typeof(IFormatConditions);
                return _type;
            }
        }

        #endregion

		#region Ctor

		/// <summary>
		/// Stub Ctor, not indented to use
		/// </summary>
		public IFormatConditions() : base()
		{

		}

		#endregion

        #region Properties

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Application Application
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", typeof(NetOffice.ExcelApi.Application));
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.Enums.XlCreator Creator
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteEnumPropertyGet<NetOffice.ExcelApi.Enums.XlCreator>(this, "Creator");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// Unknown COM Proxy
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16), ProxyResult]
        public virtual object Parent
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteReferencePropertyGet(this, "Parent");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Count
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteInt32PropertyGet(this, "Count");
            }
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// Get
        /// </summary>
        /// <param name="index">object index</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        [NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
        public virtual NetOffice.ExcelApi.FormatCondition this[object index]
        {
            get
            {
                return InvokerService.InvokeInternal.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.FormatCondition>(this, "_Default", typeof(NetOffice.ExcelApi.FormatCondition), index);
            }
        }

        #endregion

        #region Methods

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", typeof(NetOffice.ExcelApi.FormatCondition), type, _operator, formula1, formula2);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        /// <param name="textOperator">optional object textOperator</param>
        /// <param name="dateOperator">optional object dateOperator</param>
        /// <param name="scopeType">optional object scopeType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator, object dateOperator, object scopeType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Add", new object[] { type, _operator, formula1, formula2, _string, textOperator, dateOperator, scopeType });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", typeof(NetOffice.ExcelApi.FormatCondition), type);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", typeof(NetOffice.ExcelApi.FormatCondition), type, _operator);
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        [CustomMethod]
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1)
        {
            return InvokerService.InvokeInternal.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", typeof(NetOffice.ExcelApi.FormatCondition), type, _operator, formula1);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Add", new object[] { type, _operator, formula1, formula2, _string });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        /// <param name="textOperator">optional object textOperator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Add", new object[] { type, _operator, formula1, formula2, _string, textOperator });
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
        /// <param name="_operator">optional object operator</param>
        /// <param name="formula1">optional object formula1</param>
        /// <param name="formula2">optional object formula2</param>
        /// <param name="_string">optional object string</param>
        /// <param name="textOperator">optional object textOperator</param>
        /// <param name="dateOperator">optional object dateOperator</param>
        [CustomMethod]
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator, object dateOperator)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "Add", new object[] { type, _operator, formula1, formula2, _string, textOperator, dateOperator });
        }

        /// <summary>
        /// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual Int32 Delete()
        {
            return InvokerService.InvokeInternal.ExecuteInt32MethodGet(this, "Delete");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        /// <param name="colorScaleType">Int32 colorScaleType</param>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddColorScale(Int32 colorScaleType)
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddColorScale", colorScaleType);
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddDatabar()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddDatabar");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddIconSetCondition()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddIconSetCondition");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddTop10()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddTop10");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddAboveAverage()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddAboveAverage");
        }

        /// <summary>
        /// SupportByVersion Excel 12, 14, 15, 16
        /// </summary>
        [SupportByVersion("Excel", 12, 14, 15, 16)]
        public virtual object AddUniqueValues()
        {
            return InvokerService.InvokeInternal.ExecuteVariantMethodGet(this, "AddUniqueValues");
        }

        #endregion

        #region IEnumerableProvider<NetOffice.ExcelApi.FormatCondition>

        ICOMObject IEnumerableProvider<NetOffice.ExcelApi.FormatCondition>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this, false);
        }

        IEnumerable IEnumerableProvider<NetOffice.ExcelApi.FormatCondition>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, false);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.FormatCondition>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public virtual IEnumerator<NetOffice.ExcelApi.FormatCondition> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (NetOffice.ExcelApi.FormatCondition item in innerEnumerator)
                yield return item;
        }

        #endregion

        #region IEnumerable

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        IEnumerator NetRuntimeSystem.Collections.IEnumerable.GetEnumerator()
        {
            return NetOffice.Utils.GetProxyEnumeratorAsProperty(this, false);
        }

        #endregion

        #pragma warning restore
    }
}

