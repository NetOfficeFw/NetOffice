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
	/// DispatchInterface FormatConditions 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions"/> </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property), HasIndexProperty(IndexInvoke.Property, "_Default")]
	public class FormatConditions : COMObject, IEnumerableProvider<object>
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
                    _type = typeof(FormatConditions);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public FormatConditions(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public FormatConditions(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public FormatConditions(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Application"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.Application Application
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.Application>(this, "Application", NetOffice.ExcelApi.Application.LateBindingApiWrapperType);
			}
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Creator"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Parent"/> </remarks>
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
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Count"/> </remarks>
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
		public NetOffice.ExcelApi.FormatCondition this[object index]
		{
			get
			{
				return Factory.ExecuteKnownReferencePropertyGet<NetOffice.ExcelApi.FormatCondition>(this, "_Default", NetOffice.ExcelApi.FormatCondition.LateBindingApiWrapperType, index);
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", NetOffice.ExcelApi.FormatCondition.LateBindingApiWrapperType, type, _operator, formula1, formula2);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		/// <param name="_string">optional object string</param>
		/// <param name="textOperator">optional object textOperator</param>
		/// <param name="dateOperator">optional object dateOperator</param>
		/// <param name="scopeType">optional object scopeType</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator, object dateOperator, object scopeType)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", new object[]{ type, _operator, formula1, formula2, _string, textOperator, dateOperator, scopeType });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", NetOffice.ExcelApi.FormatCondition.LateBindingApiWrapperType, type);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", NetOffice.ExcelApi.FormatCondition.LateBindingApiWrapperType, type, _operator);
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public NetOffice.ExcelApi.FormatCondition Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1)
		{
			return Factory.ExecuteKnownReferenceMethodGet<NetOffice.ExcelApi.FormatCondition>(this, "Add", NetOffice.ExcelApi.FormatCondition.LateBindingApiWrapperType, type, _operator, formula1);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		/// <param name="_string">optional object string</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", new object[]{ type, _operator, formula1, formula2, _string });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		/// <param name="_string">optional object string</param>
		/// <param name="textOperator">optional object textOperator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", new object[]{ type, _operator, formula1, formula2, _string, textOperator });
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Add"/> </remarks>
		/// <param name="type">NetOffice.ExcelApi.Enums.XlFormatConditionType type</param>
		/// <param name="_operator">optional object operator</param>
		/// <param name="formula1">optional object formula1</param>
		/// <param name="formula2">optional object formula2</param>
		/// <param name="_string">optional object string</param>
		/// <param name="textOperator">optional object textOperator</param>
		/// <param name="dateOperator">optional object dateOperator</param>
		[CustomMethod]
		[SupportByVersion("Excel", 12,14,15,16)]
		public object Add(NetOffice.ExcelApi.Enums.XlFormatConditionType type, object _operator, object formula1, object formula2, object _string, object textOperator, object dateOperator)
		{
			return Factory.ExecuteVariantMethodGet(this, "Add", new object[]{ type, _operator, formula1, formula2, _string, textOperator, dateOperator });
		}

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.Delete"/> </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		public void Delete()
		{
			 Factory.ExecuteMethod(this, "Delete");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddColorScale"/> </remarks>
		/// <param name="colorScaleType">Int32 colorScaleType</param>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddColorScale(Int32 colorScaleType)
		{
			return Factory.ExecuteVariantMethodGet(this, "AddColorScale", colorScaleType);
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddDatabar"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddDatabar()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddDatabar");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddIconSetCondition"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddIconSetCondition()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddIconSetCondition");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddTop10"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddTop10()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddTop10");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddAboveAverage"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddAboveAverage()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddAboveAverage");
		}

		/// <summary>
		/// SupportByVersion Excel 12, 14, 15, 16
		/// </summary>
		/// <remarks> Docs: <see href="https://docs.microsoft.com/en-us/office/vba/api/Excel.FormatConditions.AddUniqueValues"/> </remarks>
		[SupportByVersion("Excel", 12,14,15,16)]
		public object AddUniqueValues()
		{
			return Factory.ExecuteVariantMethodGet(this, "AddUniqueValues");
		}

        #endregion

        #region IEnumerableProvider<object>

        ICOMObject IEnumerableProvider<object>.GetComObjectEnumerator(ICOMObject parent)
        {
            return NetOffice.Utils.GetComObjectEnumeratorAsProperty(parent, this);
        }

        IEnumerable IEnumerableProvider<object>.FetchVariantComObjectEnumerator(ICOMObject parent, ICOMObject enumerator)
        {
            return NetOffice.Utils.FetchVariantComObjectEnumerator(parent, enumerator, true);
        }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.FormatCondition>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        public IEnumerator<object> GetEnumerator()
        {
            NetRuntimeSystem.Collections.IEnumerable innerEnumerator = (this as NetRuntimeSystem.Collections.IEnumerable);
            foreach (object item in innerEnumerator)
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
			return NetOffice.Utils.GetProxyEnumeratorAsProperty(this);
		}

		#endregion

		#pragma warning restore
	}
}