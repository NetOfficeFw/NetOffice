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
	/// DispatchInterface Names 
	/// SupportByVersion Excel, 9,10,11,12,14,15,16
	/// </summary>
	/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff841280.aspx </remarks>
	[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
	[EntityType(EntityType.IsDispatchInterface), Enumerator(Enumerator.Reference, EnumeratorInvoke.Property, "Excel", 9, 10, 11, 12, 14, 15, 16), HasIndexProperty(IndexInvoke.Method, "_Default")]
	[TypeId("000208B8-0000-0000-C000-000000000046")]
	public interface Names : ICOMObject, IEnumerableProvider<NetOffice.ExcelApi.Name>
	{
		#region Properties

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff195100.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Application Application { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff823118.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Enums.XlCreator Creator { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// Unknown COM Proxy
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff196647.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16), ProxyResult]
		object Parent { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Get
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff840866.aspx </remarks>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		Int32 Count { get; }

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		/// <param name="refersToR1C1Local">optional object refersToR1C1Local</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1, object refersToR1C1Local);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add();

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <remarks> MSDN Online: http://msdn.microsoft.com/en-us/en-us/library/office/ff835300.aspx </remarks>
		/// <param name="name">optional object name</param>
		/// <param name="refersTo">optional object refersTo</param>
		/// <param name="visible">optional object visible</param>
		/// <param name="macroType">optional object macroType</param>
		/// <param name="shortcutKey">optional object shortcutKey</param>
		/// <param name="category">optional object category</param>
		/// <param name="nameLocal">optional object nameLocal</param>
		/// <param name="refersToLocal">optional object refersToLocal</param>
		/// <param name="categoryLocal">optional object categoryLocal</param>
		/// <param name="refersToR1C1">optional object refersToR1C1</param>
		[CustomMethod]
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		NetOffice.ExcelApi.Name Add(object name, object refersTo, object visible, object macroType, object shortcutKey, object category, object nameLocal, object refersToLocal, object categoryLocal, object refersToR1C1);

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Custom Indexer
		/// </summary>
		/// <param name="index">object index</param>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.ExcelApi.Name this[object index] { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// Custom Indexer
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="indexLocal">optional object indexLocal</param>
		[SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty, CustomIndexer]
		NetOffice.ExcelApi.Name this[object index, object indexLocal] { get; }

		/// <summary>
		/// SupportByVersion Excel 9, 10, 11, 12, 14, 15, 16
		/// </summary>
		/// <param name="index">optional object index</param>
		/// <param name="indexLocal">optional object indexLocal</param>
		/// <param name="refersTo">optional object refersTo</param>
		[SupportByVersion("Excel", 9,10,11,12,14,15,16)]
		[NetRuntimeSystem.Runtime.CompilerServices.IndexerName("Item"), IndexProperty]
		NetOffice.ExcelApi.Name this[object index, object indexLocal, object refersTo] { get; }

        #endregion

        #region IEnumerable<NetOffice.ExcelApi.Name>

        /// <summary>
        /// SupportByVersion Excel, 9,10,11,12,14,15,16
        /// </summary>
        [SupportByVersion("Excel", 9, 10, 11, 12, 14, 15, 16)]
        new IEnumerator<NetOffice.ExcelApi.Name> GetEnumerator();

        #endregion
    }
}
